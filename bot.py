# -*- coding: utf-8 -*-
"""
Telegram-бот: принимает файлы филиалов и отдаёт готовый
«Ежедневный отчёт по просрочке».

Работает на long polling — белый IP, домен и сертификаты не нужны.
Из внешних библиотек только requests и openpyxl.
"""

from __future__ import annotations

import configparser
import json
import logging
import os
import shutil
import sys
import time
import traceback
from datetime import datetime, timedelta

import requests

import engine as e

ЗДЕСЬ = os.path.dirname(os.path.abspath(__file__))
КОНФИГ = os.path.join(ЗДЕСЬ, "config.ini")
ДАННЫЕ = os.path.join(ЗДЕСЬ, "data")
ШАБЛОН = os.path.join(ЗДЕСЬ, "template", "Ежедневный_отчёт_по_просрочке.xlsx")
КРЕДИТЫ_ЗАПАС = os.path.join(ЗДЕСЬ, "template", "Активные_кредиты.xlsx")
КРЕДИТЫ = os.path.join(ДАННЫЕ, "Активные_кредиты.xlsx")
ВЧЕРА = os.path.join(ДАННЫЕ, "вчера.json")
ЖУРНАЛ = os.path.join(ДАННЫЕ, "журнал.json")
ИМЯ_ОТЧЁТА = "Ежедневный_отчёт_по_просрочке.xlsx"
ИМЯ_СПИСКА = "Kotarmadi_va_ishlanmagan.xlsx"

MAX_МБ = 45
ТАЙМАУТ = 50
ДНЕЙ_В_ИСТОРИИ = 14
МАКС_ЗАПИСЕЙ = 5000
log = logging.getLogger("bot")


# ------------------------------------------------------------------ настройки

def прочитать_конфиг():
    if not os.path.exists(КОНФИГ):
        sys.exit("Нет config.ini рядом с ботом. Скопируйте config.example.ini "
                 "в config.ini и заполните.")
    cp = configparser.ConfigParser()
    cp.read(КОНФИГ, encoding="utf-8")
    токен = cp.get("bot", "token", fallback="").strip()
    if not токен or токен.startswith("СЮДА"):
        sys.exit("В config.ini не указан token бота.")
    сырые = cp.get("bot", "allowed_ids", fallback="").replace(",", " ").split()
    разрешённые = {int(x) for x in сырые if x.strip().lstrip("-").isdigit()}
    статус = cp.get("format", "status_column", fallback="E").strip().upper()
    коммент = cp.get("format", "comment_column", fallback="F").strip().upper()
    return токен, разрешённые, статус, коммент


# ------------------------------------------------------------------ хранение дня

def папка_дня(дата=None):
    д = (дата or datetime.now()).strftime("%Y-%m-%d")
    п = os.path.join(ДАННЫЕ, д)
    os.makedirs(п, exist_ok=True)
    return п


def файлы_дня(дата=None):
    """Только присланные файлы: собранные отчёт и список лежат рядом и входными не считаются."""
    п = папка_дня(дата)
    return sorted(os.path.join(п, f) for f in os.listdir(п)
                  if f.lower().endswith(".xlsx") and f not in (ИМЯ_ОТЧЁТА, ИМЯ_СПИСКА))


def свободный_путь(папка, имя):
    """Пачка файлов приходит в одну секунду; одинаковые имена не должны затирать
    друг друга — иначе бот ответит «принял» на все, а на диске останется один."""
    основа, расш = os.path.splitext(имя)
    метка = int(time.time())
    кандидат = os.path.join(папка, "%d_%s%s" % (метка, основа, расш))
    n = 2
    while os.path.exists(кандидат):
        кандидат = os.path.join(папка, "%d_%s(%d)%s" % (метка, основа, n, расш))
        n += 1
    return кандидат


def убрать_файлы_дня():
    """Отчёт выдан — присланное больше не нужно. Сам отчёт остаётся как след."""
    убрано = 0
    for путь in файлы_дня():
        try:
            os.remove(путь)
            убрано += 1
        except OSError:
            log.exception("не удалился файл %s", путь)
    return убрано


def отчёт_собран(дата=None):
    """Время сборки сегодняшнего отчёта или None. По нему видно: день уже закрыт."""
    путь = os.path.join(папка_дня(дата), ИМЯ_ОТЧЁТА)
    if not os.path.exists(путь):
        return None
    return datetime.fromtimestamp(os.path.getmtime(путь))


def очистить_день():
    п = папка_дня()
    shutil.rmtree(п, ignore_errors=True)
    os.makedirs(п, exist_ok=True)


def сохранить_вчерашние(объединение):
    """Снимок просрочки по кредитам — завтра по нему посчитаем оплаты."""
    снимок = {"дата": datetime.now().strftime("%Y-%m-%d"),
              "просрочка": {к: d["просрочка"] for к, d in объединение.items()
                            if not к.startswith(e.БЕЗ_КОДА)}}   # «?12» завтра ничего не значит
    with open(ВЧЕРА, "w", encoding="utf-8") as f:
        json.dump(снимок, f)


# ------------------------------------------------------------------ журнал загрузок

def прочитать_журнал():
    """Список записей о каждом принятом файле. Переживает «Очистить день»."""
    if not os.path.exists(ЖУРНАЛ):
        return []
    try:
        with open(ЖУРНАЛ, encoding="utf-8") as f:
            записи = json.load(f)
        return записи if isinstance(записи, list) else []
    except Exception:
        log.exception("журнал не прочитался")
        return []


def записать_в_журнал(запись):
    журнал = прочитать_журнал()
    журнал.append(запись)
    del журнал[:-МАКС_ЗАПИСЕЙ]
    врем = ЖУРНАЛ + ".tmp"
    try:
        with open(врем, "w", encoding="utf-8") as f:
            json.dump(журнал, f, ensure_ascii=False, indent=1)
        os.replace(врем, ЖУРНАЛ)          # подмена целиком — журнал не побьётся
    except Exception:
        log.exception("журнал не записался")


def кто_прислал(отправитель):
    """«Иван Петров (@ivan)» — чтобы в истории было видно человека, а не номер."""
    отправитель = отправитель or {}
    фио = " ".join(x for x in (отправитель.get("first_name"),
                               отправитель.get("last_name")) if x)
    ник = отправитель.get("username")
    if ник:
        фио = ("%s (@%s)" % (фио, ник)).strip()
    return фио or str(отправитель.get("id", "?"))


def запись_журнала(отправитель, имя, тип, размер, строк, заполнено, области, сохранён):
    сейчас = datetime.now()
    return {"дата": сейчас.strftime("%Y-%m-%d"),
            "время": сейчас.strftime("%H:%M:%S"),
            "кто": (отправитель or {}).get("id"),
            "имя": кто_прислал(отправитель),
            "файл": имя,
            "сохранён": сохранён,
            "тип": тип,
            "строк": строк,
            "заполнено": заполнено,
            "области": области,
            "размер": размер}


def журнал_за_день(дата=None):
    """{сохранённое имя файла: запись} — чтобы подтянуть время и автора."""
    ключ = (дата or datetime.now()).strftime("%Y-%m-%d")
    return {з["сохранён"]: з for з in прочитать_журнал()
            if з.get("дата") == ключ and з.get("сохранён")}


def откуда_файл(сохранённое, журнал_дня):
    """Настоящее имя файла и «когда · кто». Для файлов до журнала — время из префикса."""
    з = журнал_дня.get(сохранённое)
    if з:
        return з["файл"], "%s · %s" % (з["время"][:5], з["имя"])
    части = сохранённое.split("_", 1)
    if len(части) == 2 and части[0].isdigit():
        return части[1], datetime.fromtimestamp(int(части[0])).strftime("%H:%M")
    return сохранённое, "—"


def взять_вчерашние():
    if not os.path.exists(ВЧЕРА):
        return None, None
    try:
        with open(ВЧЕРА, encoding="utf-8") as f:
            с = json.load(f)
        if с.get("дата") == datetime.now().strftime("%Y-%m-%d"):
            return None, None          # снимок сегодняшний, разницы нет
        return с["просрочка"], с["дата"]
    except Exception:
        return None, None


# ------------------------------------------------------------------ Telegram

КНОПКА = lambda текст, код: {"text": текст, "callback_data": код}


def меню(главное=True):
    if главное:
        ряды = [[КНОПКА("📊 Hisobotni yig'ish", "отчёт")],
                [КНОПКА("📁 Nima yuklangan", "состав"),
                 КНОПКА("🗓 Yuklamalar tarixi", "история")],
                [КНОПКА("🗑 Kunni tozalash", "очистить"), КНОПКА("❓ Yordam", "помощь")]]
    else:
        ряды = [[КНОПКА("📊 Hisobotni yig'ish", "отчёт"), КНОПКА("📁 Nima yuklangan", "состав")]]
    return {"inline_keyboard": ряды}


МЕНЮ_ПОСЛЕ_ФАЙЛА = {"inline_keyboard": [
    [КНОПКА("📊 Hisobotni hozir yig'ish", "отчёт")],
    [КНОПКА("📁 Nima yuklangan", "состав"), КНОПКА("🗑 Kunni tozalash", "очистить")]]}

МЕНЮ_ПОСЛЕ_ОТЧЁТА = {"inline_keyboard": [
    [КНОПКА("🗓 Yuklamalar tarixi", "история"), КНОПКА("🏠 Menyu", "меню")]]}

ПОДТВЕРДИТЬ_ОЧИСТКУ = {"inline_keyboard": [
    [КНОПКА("Ha, tozalash", "очистить_да"), КНОПКА("Bekor qilish", "меню")]]}


class Телеграм:
    def __init__(self, токен):
        self.база = "https://api.telegram.org/bot%s" % токен
        self.файлы = "https://api.telegram.org/file/bot%s" % токен
        self.с = requests.Session()

    def _вызов(self, метод, **kw):
        r = self.с.post("%s/%s" % (self.база, метод), timeout=90, **kw)
        r.raise_for_status()
        д = r.json()
        if not д.get("ok"):
            raise RuntimeError("Telegram: %s" % д)
        return д["result"]

    def я(self):
        return self._вызов("getMe")

    def обновления(self, offset):
        r = self.с.get("%s/getUpdates" % self.база,
                       params={"offset": offset, "timeout": ТАЙМАУТ,
                               "allowed_updates": '["message","callback_query"]'},
                       timeout=ТАЙМАУТ + 20)
        r.raise_for_status()
        д = r.json()
        return д.get("result", []) if д.get("ok") else []

    def написать(self, chat_id, текст, кнопки=None, ответ_на=None):
        д = {"chat_id": chat_id, "text": текст[:4000], "disable_web_page_preview": True}
        if кнопки:
            д["reply_markup"] = json.dumps(кнопки, ensure_ascii=False)
        if ответ_на:
            д["reply_to_message_id"] = ответ_на
            д["allow_sending_without_reply"] = True
        try:
            return self._вызов("sendMessage", data=д)
        except Exception:
            log.exception("не отправилось сообщение")

    def ответ_на_кнопку(self, callback_id, текст=""):
        try:
            self._вызов("answerCallbackQuery",
                        data={"callback_query_id": callback_id, "text": текст[:190]})
        except Exception:
            pass

    def действие(self, chat_id, что="upload_document"):
        try:
            self._вызов("sendChatAction", data={"chat_id": chat_id, "action": что})
        except Exception:
            pass

    def отправить_файл(self, chat_id, путь, имя, подпись, кнопки=None):
        д = {"chat_id": chat_id, "caption": подпись[:1000]}
        if кнопки:
            д["reply_markup"] = json.dumps(кнопки, ensure_ascii=False)
        with open(путь, "rb") as f:
            return self._вызов("sendDocument", data=д, files={"document": (имя, f)})

    def скачать(self, file_id, куда):
        инфо = self._вызов("getFile", data={"file_id": file_id})
        r = self.с.get("%s/%s" % (self.файлы, инфо["file_path"]), timeout=300, stream=True)
        r.raise_for_status()
        with open(куда, "wb") as f:
            for кусок in r.iter_content(1 << 16):
                f.write(кусок)
        return куда


# ------------------------------------------------------------------ работа

ПРИВЕТ = (
    "Bot «Muddati o'tgan qarzlar bo'yicha kunlik hisobot»ni yig'adi.\n\n"
    "Filial fayllarini hujjat sifatida yuboring — bittalab yoki bir nechtasini birdan. "
    "Hammasi kelgach, «Hisobotni yig'ish» tugmasini bosing.\n\n"
    "Hisobot bilan birga ikkinchi fayl ham keladi — kimga qayta qo'ng'iroq qilish "
    "kerak, ikki varaqda: hali ishlanmaganlar va telefonni ko'tarmaganlar. "
    "Filial bo'yicha filtr bilan.\n\n"
    "⚠️ Hisobot bir marta yig'iladi: undan keyin yuklangan fayllar tozalanadi. "
    "Yangi hisobot kerak bo'lsa, barcha fayllarni qaytadan yuboring — "
    "shuning uchun tugmani hamma fayllar kelgandan keyin bosing.\n\n"
    "Filial faylida qo'ng'iroq holati E ustunidan, izoh F ustunidan olinadi. "
    "Agar holat «To'ladi» bo'lsa, to'lov summasi izohga yoziladi — bot uni hisoblaydi.\n\n"
    "Ikki haftada bir marta «Faol kreditlar» yangi yuklamasini yuboring — "
    "bot uni eslab qoladi va portfelni shu bo'yicha hisoblaydi.\n\n"
    "«🗓 Yuklamalar tarixi» — kim qaysi kuni qaysi faylni yuborganini ko'rsatadi "
    "va yuklama bo'lmagan kunlarni belgilaydi."
)


def разобрать_день(статус_кол, коммент_кол):
    """Читает все файлы за сегодня. Возвращает (файлы, ошибки, сведения)."""
    файлы, ошибки, сведения = {}, [], {}
    for путь in файлы_дня():
        имя = os.path.basename(путь)
        try:
            строки, свед = e.читать_просрочку(путь, статус_кол, коммент_кол)
            файлы[имя] = строки
            сведения[имя] = свед
        except e.ОшибкаФайла as ош:
            ошибки.append((имя, str(ош)))
        except Exception:
            log.exception("сбой разбора %s", имя)
            ошибки.append((имя, "faylni o'qiy olmadim"))
    return файлы, ошибки, сведения


def портфель():
    источник = КРЕДИТЫ if os.path.exists(КРЕДИТЫ) else КРЕДИТЫ_ЗАПАС
    сумма, штук = e.читать_кредиты(источник)
    return сумма, штук, (источник == КРЕДИТЫ)


def сделать_отчёт(тг, chat_id, статус_кол, коммент_кол, отправитель=None):
    файлы, ошибки, сведения = разобрать_день(статус_кол, коммент_кол)
    if not файлы:
        когда = отчёт_собран()
        if когда:
            тг.написать(chat_id,
                        "✅ Bugungi hisobot yig'ilgan\n"
                        "⏰ soat %s\n\n"
                        "Shundan keyin fayllar tozalandi. Yangi hisobot uchun "
                        "barcha filial fayllarini qaytadan yuboring."
                        % когда.strftime("%H:%M"), меню())
        else:
            тг.написать(chat_id,
                        "📭 Bugun uchun fayl yo'q\n\n"
                        "Filial fayllarini hujjat sifatida yuboring — "
                        "hammasi kelgach hisobotni yig'aman.", меню())
        return

    дубли = e.дубликаты(файлы)
    for имя, оригинал in дубли:
        файлы.pop(имя, None)

    вчера_просрочка, _ = взять_вчерашние()
    итог = e.свести(файлы, вчера_просрочка)
    пор_сум, пор_шт, свой = портфель()

    тг.действие(chat_id)
    куда = os.path.join(папка_дня(), ИМЯ_ОТЧЁТА)
    e.собрать_отчёт(итог, пор_сум, пор_шт, ШАБЛОН, куда)
    сохранить_вчерашние(итог["объединение"])

    заметки = ["📁 Fayllar: %d ta" % len(файлы)]
    if дубли:
        заметки.append("♻️ Nusxa (bir marta hisoblandi): %s"
                       % ", ".join(a for a, _ in дубли))
    if ошибки:
        заметки.append("🚫 O'qilmadi: %s" % "; ".join(и for и, _ in ошибки))
    if not свой:
        заметки.append("ℹ️ Portfel — o'rnatishdagi kreditlar yuklamasi bo'yicha")
    заметки.append("⚠️ Hisobot bir marta yig'iladi — fayllar tozalandi, "
                   "keyingisi uchun qaytadan yuboring")

    подпись = e.текст_сводки(итог, пор_сум, пор_шт, заметки)
    имя_файла = "Muddati_otgan_qarzlar_hisoboti_%s.xlsx" % datetime.now().strftime("%d.%m.%Y")
    тг.отправить_файл(chat_id, куда, имя_файла, подпись, МЕНЮ_ПОСЛЕ_ОТЧЁТА)
    try:
        отправить_список(тг, chat_id, итог)
    except Exception:
        # отчёт уже у человека — из-за списка его не повторяем
        log.exception("список недозвона не собрался")
        тг.написать(chat_id, "⚠️ Ko'tarmagan va ishlanmaganlar ro'yxatini yubora olmadim — "
                             "tafsilotlar bot jurnalida.")

    # только после успешной отправки: если Telegram не принял файл, загрузки уцелеют
    убрано = убрать_файлы_дня()
    записать_в_журнал(запись_журнала(отправитель, имя_файла, "отчёт", os.path.getsize(куда),
                                     итог["всего_строк"], итог["обработано"], [], ""))
    log.info("отчёт выдан, файлов освобождено: %d", убрано)


def отправить_список(тг, chat_id, итог):
    """Вслед за отчётом — кому не дозвонились и до кого не дошли. Только сейчас:
    после отчёта присланные файлы удаляются, и собрать список будет не из чего."""
    список = e.кого_не_достали(итог)
    if not список:
        тг.написать(chat_id, "📞 Bugun barcha kreditlar bo'yicha aloqa bo'ldi — "
                             "ko'tarmagan va ishlanmagan yo'q.")
        return
    куда = os.path.join(папка_дня(), ИМЯ_СПИСКА)
    e.собрать_список(список, куда)
    имя = "Kotarmadi_va_ishlanmagan_%s.xlsx" % datetime.now().strftime("%d.%m.%Y")
    тг.отправить_файл(chat_id, куда, имя, e.текст_списка(список))


def показать_состав(тг, chat_id, статус_кол, коммент_кол):
    файлы, ошибки, сведения = разобрать_день(статус_кол, коммент_кол)
    if not файлы and not ошибки:
        тг.написать(chat_id, "Bugun uchun hali fayllar yo'q.", меню())
        return
    журнал_дня = журнал_за_день()
    строки = ["📁 YUKLANGAN FAYLLAR", datetime.now().strftime("%d.%m.%Y"), ""]
    всего_пометок = 0
    for n, (имя, свед) in enumerate(сведения.items(), 1):
        обл = ", ".join(e._кратко(о) for о in свед["области"]) or "aniqlanmadi"
        настоящее, когда = откуда_файл(имя, журнал_дня)
        всего_пометок += свед["заполнено"]
        строки.append("%d. %s" % (n, настоящее))
        строки.append("   ⏰ %s" % когда)
        строки.append("   📋 %s qator · 📍 %s" % (e.штук(свед["заполнено"]), обл))
    дубли = e.дубликаты(файлы)
    if дубли:
        строки += ["", "♻️ Nusxa: " + ", ".join(
            откуда_файл(a, журнал_дня)[0] for a, _ in дубли)]
    for имя, ош in ошибки:
        строки.append("🚫 %s — %s" % (откуда_файл(имя, журнал_дня)[0], ош))
    строки += ["", "Jami: %d fayl · %s belgi" % (len(сведения), e.штук(всего_пометок))]
    тг.написать(chat_id, "\n".join(строки), меню())


def показать_историю(тг, chat_id):
    """Ежедневная загрузка за две недели — видно и дни, когда никто не прислал."""
    по_дням = {}
    for з in прочитать_журнал():
        по_дням.setdefault(з.get("дата"), []).append(з)

    строки = ["🗓 Oxirgi %d kundagi yuklamalar:" % ДНЕЙ_В_ИСТОРИИ, ""]
    сегодня = datetime.now()
    было_всего = 0
    for n in range(ДНЕЙ_В_ИСТОРИИ):
        день = сегодня - timedelta(days=n)
        записи = по_дням.get(день.strftime("%Y-%m-%d"), [])
        подпись = день.strftime("%d.%m")
        if n == 0:
            подпись += " (bugun)"
        загрузки = [з for з in записи if з.get("тип") in ("просрочка", "кредиты")]
        if not загрузки:
            строки.append("%s — ❌ yuklama yo'q" % подпись)
        else:
            было_всего += len(загрузки)
            филиалов = sum(1 for з in загрузки if з.get("тип") == "просрочка")
            итоги = ["%d fayl" % филиалов] if филиалов else []
            if len(загрузки) > филиалов:
                итоги.append("kreditlar yuklamasi")
            строки.append("%s — ✅ %s" % (подпись, " + ".join(итоги)))
        for з in записи:
            что = ("📊 hisobot yig'ildi" if з.get("тип") == "отчёт"
                   else з.get("файл", "?"))
            строки.append("   %s %s — %s" % (з.get("время", "")[:5],
                                             з.get("имя", "?"), что))
    if not было_всего:
        строки += ["", "Jurnal bo'sh: bot yangi fayllarni shu yerda qayd qila boshlaydi."]
    тг.написать(chat_id, "\n".join(строки), меню())


def принять_документ(тг, сообщение, chat_id, mid, статус_кол, коммент_кол):
    док = сообщение["document"]
    имя = док.get("file_name") or "файл.xlsx"
    размер = док.get("file_size") or 0

    if not имя.lower().endswith((".xlsx", ".xlsm")):
        тг.написать(chat_id, "Excel fayl (.xlsx) kerak. «%s» yuborildi.\n"
                             "Agar rasm sifatida yuborgan bo'lsangiz — hujjat sifatida yuboring."
                    % имя, меню(False), mid)
        return
    if размер > MAX_МБ * 1024 * 1024:
        тг.написать(chat_id,
                    "Fayl juda katta (%.1f MB). Telegram botlarga %d MB gacha beradi.\n"
                    "Odatda bu «shishgan» ish diapazoni: faylni oching, Ctrl+End bosing, "
                    "ortiqcha qator va ustunlarni o'chiring, saqlang."
                    % (размер / 1048576, MAX_МБ), меню(False), mid)
        return

    тг.действие(chat_id, "typing")
    временный = os.path.join(ДАННЫЕ, "вход_%d.xlsx" % int(time.time() * 1000))
    try:
        тг.скачать(док["file_id"], временный)
    except Exception:
        log.exception("не скачался файл")
        тг.написать(chat_id, "Faylni yuklab ololmadim. Yana bir marta yuboring.",
                    меню(False), mid)
        return

    try:
        тип = e.определить_тип(временный)
        if тип == "кредиты":
            shutil.move(временный, КРЕДИТЫ)
            сум, шт = e.читать_кредиты(КРЕДИТЫ)
            записать_в_журнал(запись_журнала(
                сообщение.get("from"), имя, "кредиты", размер,
                sum(шт.values()), 0, sorted(сум), ""))
            тг.написать(chat_id,
                        "Kreditlar yuklamasini qabul qildim va eslab qoldim.\n"
                        "Viloyatlar: %d · portfel %.2f mlrd so'm.\n\n"
                        "Endi filial fayllarini yuboring."
                        % (len(сум), sum(сум.values()) / 1_000_000_000), меню(), mid)
            return

        if тип != "просрочка":
            тг.написать(chat_id,
                        "Bu qanday fayl ekanini tushunmadim.\n\n"
                        "Ikkitadan birini kutyapman:\n"
                        "• muddati o'tgan qarzlar yuklamasi — «Итого просрочка» ustuni bilan\n"
                        "• «Faol kreditlar» — «Филиал» va «Общая задолженность» ustunlari bilan",
                        меню(False), mid)
            return

        строки, свед = e.читать_просрочку(временный, статус_кол, коммент_кол)
        цель = свободный_путь(папка_дня(), имя)
        shutil.move(временный, цель)
        записать_в_журнал(запись_журнала(
            сообщение.get("from"), имя, "просрочка", размер,
            свед["строк"], свед["заполнено"], свед["области"], os.path.basename(цель)))

        обл = ", ".join(e._кратко(о) for о in свед["области"]) or "aniqlanmadi"
        оплат = sum(1 for d in строки.values() if d["статус"] == "Оплачено")
        сумма_оплат = sum(d["оплата"] for d in строки.values() if d["статус"] == "Оплачено")
        текст = ["✅ Qabul qilindi", имя, "",
                 "📋 To'ldirilgan: %s / %s qator"
                 % (e.штук(свед["заполнено"]), e.штук(свед["строк"])),
                 "📍 Viloyat: %s" % обл]
        if оплат:
            текст.append("💰 To'lovlar: %d ta → %.1f mln so'm"
                         % (оплат, сумма_оплат / 1_000_000))
        if свед["заполнено"] == 0:
            текст.append("⚠️ Bu faylda birorta ham belgi yo'q")
        if свед.get("без_кода"):
            # отчёт найдёт настоящий код по ФИО в других файлах, но исправить лучше в самом файле
            текст.append("⚠️ «Код клиента» buzilgan (%d qator): %s — iltimos, faylda tuzating"
                         % (len(свед["без_кода"]), "; ".join(свед["без_кода"][:3])))
        текст += ["", "📁 Bugun jami: %d fayl" % len(файлы_дня())]
        тг.написать(chat_id, "\n".join(текст), МЕНЮ_ПОСЛЕ_ФАЙЛА, mid)

    except e.ОшибкаФайла as ош:
        тг.написать(chat_id, "Fayl bilan nimadir noto'g'ri:\n\n%s" % ош, меню(False), mid)
    except Exception:
        log.exception("сбой обработки файла")
        тг.написать(chat_id, "Faylni qayta ishlay olmadim — ichki xato. "
                             "Tafsilotlar bot jurnalida.", меню(False), mid)
    finally:
        if os.path.exists(временный):
            try:
                os.remove(временный)
            except OSError:
                pass


def обработать_кнопку(тг, запрос, разрешённые, статус_кол, коммент_кол):
    chat_id = запрос["message"]["chat"]["id"]
    кто = запрос["from"]["id"]
    код = запрос.get("data", "")
    тг.ответ_на_кнопку(запрос["id"])

    if разрешённые and кто not in разрешённые:
        тг.написать(chat_id, "Sizda ruxsat yo'q. Sizning ID: %s" % кто)
        return

    if код == "отчёт":
        сделать_отчёт(тг, chat_id, статус_кол, коммент_кол, запрос.get("from"))
    elif код == "состав":
        показать_состав(тг, chat_id, статус_кол, коммент_кол)
    elif код == "история":
        показать_историю(тг, chat_id)
    elif код == "очистить":
        тг.написать(chat_id, "Bugungi barcha fayllar o'chirilsinmi? "
                             "Hisobotni qaytadan yig'ishga to'g'ri keladi.",
                    ПОДТВЕРДИТЬ_ОЧИСТКУ)
    elif код == "очистить_да":
        очистить_день()
        тг.написать(chat_id, "Bugungi fayllar o'chirildi. Qaytadan yuboring.", меню())
    elif код == "помощь":
        тг.написать(chat_id, ПРИВЕТ, меню())
    else:
        тг.написать(chat_id, "Nimada yordam bera olaman?", меню())


def обработать_сообщение(тг, сообщение, разрешённые, статус_кол, коммент_кол):
    chat_id = сообщение["chat"]["id"]
    кто = (сообщение.get("from") or {}).get("id")
    mid = сообщение.get("message_id")

    if разрешённые and кто not in разрешённые:
        log.warning("отказ: id=%s (%s)", кто, (сообщение.get("from") or {}).get("username"))
        тг.написать(chat_id, "Sizda bu botdan foydalanish huquqi yo'q.\nSizning ID: %s — "
                             "uni administratorga bering." % кто, ответ_на=mid)
        return

    if "document" in сообщение:
        принять_документ(тг, сообщение, chat_id, mid, статус_кол, коммент_кол)
        return

    текст = (сообщение.get("text") or "").strip().lower()
    if сообщение.get("photo"):
        тг.написать(chat_id, "Bu rasm. Faylni hujjat sifatida yuboring: qisqich → «Fayl».",
                    меню(False), mid)
    elif текст.startswith(("/tarix", "/история")) or "tarix" in текст:
        показать_историю(тг, chat_id)
    elif (текст.startswith(("/hisobot", "/otchet", "/отчет"))
          or "hisobot" in текст or "отчёт" in текст):
        сделать_отчёт(тг, chat_id, статус_кол, коммент_кол, сообщение.get("from"))
    else:
        тг.написать(chat_id, ПРИВЕТ, меню(), mid)


# ------------------------------------------------------------------ цикл

def main():
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.StreamHandler(sys.stdout)])
    токен, разрешённые, статус_кол, коммент_кол = прочитать_конфиг()
    os.makedirs(ДАННЫЕ, exist_ok=True)
    тг = Телеграм(токен)

    try:
        я = тг.я()
    except requests.exceptions.HTTPError as ош:
        код = ош.response.status_code if ош.response is not None else 0
        if код == 401:
            sys.exit("Telegram не принял токен. Проверьте строку token в config.ini.")
        sys.exit("Telegram ответил ошибкой %s." % код)
    except requests.exceptions.RequestException as ош:
        sys.exit("Нет связи с Telegram (%s). Проверьте интернет на сервере." % ош)

    log.info("бот запущен: @%s | статус=%s комментарий=%s",
             я.get("username"), статус_кол, коммент_кол)
    if разрешённые:
        log.info("доступ разрешён для %d человек", len(разрешённые))
    else:
        log.warning("СПИСОК ДОСТУПА ПУСТ — бот отвечает всем!")

    offset = 0
    while True:
        try:
            for upd in тг.обновления(offset):
                offset = upd["update_id"] + 1
                try:
                    if "callback_query" in upd:
                        обработать_кнопку(тг, upd["callback_query"], разрешённые,
                                          статус_кол, коммент_кол)
                    elif "message" in upd:
                        обработать_сообщение(тг, upd["message"], разрешённые,
                                             статус_кол, коммент_кол)
                except Exception:
                    log.error("сбой на обновлении:\n%s", traceback.format_exc())
        except requests.exceptions.RequestException as ош:
            log.warning("сеть недоступна (%s), повтор через 15 с", ош)
            time.sleep(15)
        except Exception:
            log.error("непредвиденный сбой:\n%s", traceback.format_exc())
            time.sleep(10)


if __name__ == "__main__":
    main()
