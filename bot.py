import os
from datetime import datetime
import pandas as pd
import asyncio

from aiogram import Bot, Dispatcher, F
from aiogram.types import (
    Message, FSInputFile, KeyboardButton,
    ReplyKeyboardMarkup, ReplyKeyboardRemove,
    InlineKeyboardMarkup, InlineKeyboardButton,
    CallbackQuery
)
from aiogram.filters import Command
from aiogram.fsm.state import StatesGroup, State
from aiogram.fsm.context import FSMContext

# ==========================
# SOZLAMALAR
# ==========================
BOT_TOKEN = "8566383768:AAEgufQmobt2xEXTvsPWTDAkJC_xKsOBTnU"

ADMINS = [
    2132452028,
    804951151,
    467217640,
    6864825340,
    464237716,
    555555555
]

ZAYAVKA_FILE = "zayavka.xlsx"

bot = Bot(BOT_TOKEN)
dp = Dispatcher()

# ==========================
# EXCEL STRUKTURASINI TEKSHIRISH
# ==========================
COLUMNS = [
    "ID","Sana","Muassasa","Hudud",
    "AKT","Tel","Muammo",
    "Lat","Lon","UserID","Status"
]

def check_excel():
    if not os.path.exists(ZAYAVKA_FILE):
        pd.DataFrame(columns=COLUMNS).to_excel(ZAYAVKA_FILE, index=False)
        return

    df = pd.read_excel(ZAYAVKA_FILE)

    # Ustunlar mos kelmasa qayta yaratish
    if list(df.columns) != COLUMNS:
        df = pd.DataFrame(columns=COLUMNS)
        df.to_excel(ZAYAVKA_FILE, index=False)

check_excel()

# ==========================
# FSM
# ==========================
class Zayavka(StatesGroup):
    muassasa = State()
    hudud = State()
    akt = State()
    tel = State()
    muammo = State()
    lokatsiya = State()

# ==========================
# START
# ==========================
@dp.message(Command("start"))
async def start(message: Message):
    await message.answer(
        "Assalomu alaykum!\n\n"
        "🆕 Zayavka qoldirish – /zayavka\n"
        "📊 Adminlar uchun – /hisobot",
        reply_markup=ReplyKeyboardRemove()
    )

# ==========================
# ZAYAVKA BOSHLASH
# ==========================
@dp.message(Command("zayavka"))
async def z_start(message: Message, state: FSMContext):
    await state.set_state(Zayavka.muassasa)
    await message.answer("🏥 Muassasa nomini yozing:")

@dp.message(Zayavka.muassasa)
async def z_muassasa(message: Message, state: FSMContext):
    await state.update_data(muassasa=message.text)
    await state.set_state(Zayavka.hudud)
    await message.answer("📍 Hudud (tuman/shahar):")

@dp.message(Zayavka.hudud)
async def z_hudud(message: Message, state: FSMContext):
    await state.update_data(hudud=message.text)
    await state.set_state(Zayavka.akt)
    await message.answer("👤 AKT mas’ul shaxs ismi:")

@dp.message(Zayavka.akt)
async def z_akt(message: Message, state: FSMContext):
    await state.update_data(akt=message.text)
    await state.set_state(Zayavka.tel)
    await message.answer("📞 Telefon raqam:")

@dp.message(Zayavka.tel)
async def z_tel(message: Message, state: FSMContext):
    await state.update_data(tel=message.text)
    await state.set_state(Zayavka.muammo)
    await message.answer("🛠 Muammo tavsifi:")

@dp.message(Zayavka.muammo)
async def z_muammo(message: Message, state: FSMContext):
    await state.update_data(muammo=message.text)

    btn = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text="📍 Lokatsiya yuborish", request_location=True)]],
        resize_keyboard=True
    )

    await state.set_state(Zayavka.lokatsiya)
    await message.answer("📌 Lokatsiyani yuboring:", reply_markup=btn)

# ==========================
# LOKATSIYA VA YAKUN
# ==========================
@dp.message(Zayavka.lokatsiya, F.location)
async def z_finish(message: Message, state: FSMContext):

    data = await state.get_data()

    lat = message.location.latitude
    lon = message.location.longitude
    sana = datetime.now().strftime("%d.%m.%Y %H:%M")

    df = pd.read_excel(ZAYAVKA_FILE)
    new_id = len(df) + 1

    new_row = {
        "ID": new_id,
        "Sana": sana,
        "Muassasa": data["muassasa"],
        "Hudud": data["hudud"],
        "AKT": data["akt"],
        "Tel": data["tel"],
        "Muammo": data["muammo"],
        "Lat": lat,
        "Lon": lon,
        "UserID": message.from_user.id,
        "Status": "Yangi"
    }

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(ZAYAVKA_FILE, index=False)

    text = f"""
🆕 <b>YANGI ZAYAVKA #{new_id}</b>

🏥 {data["muassasa"]}
📍 {data["hudud"]}
👤 {data["akt"]}
📞 {data["tel"]}

🛠 {data["muammo"]}

📅 {sana}
Status: 🆕 Yangi
"""

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(
                text="🟡 Jarayonga olish",
                callback_data=f"progress_{new_id}"
            ),
            InlineKeyboardButton(
                text="✅ Bajarildi",
                callback_data=f"done_{new_id}"
            )
        ]
    ])

    # ADMINLARGA YUBORISH
    for a in ADMINS:
        try:
            await bot.send_message(a, text, parse_mode="HTML", reply_markup=kb)
            await bot.send_location(a, lat, lon)
        except:
            pass

    await message.answer(
        "✅ Zayavkangiz yuborildi!\nAdminlar tez ko‘rib chiqadi.",
        reply_markup=ReplyKeyboardRemove()
    )
    await state.clear()

# ==========================
# STATUS: JARAYONGA OLISH
# ==========================
@dp.callback_query(F.data.startswith("progress_"))
async def to_progress(call: CallbackQuery):

    zid = int(call.data.split("_")[1])

    df = pd.read_excel(ZAYAVKA_FILE)
    df.loc[df["ID"] == zid, "Status"] = "Jarayonda"
    df.to_excel(ZAYAVKA_FILE, index=False)

    user_id = int(df[df["ID"] == zid]["UserID"].values[0])
    try:
        await bot.send_message(
            user_id,
            f"✅ Zayavkangiz №{zid} jarayonga olindi."
        )
    except:
        pass

    await call.message.answer(f"🟡 Zayavka #{zid} – Jarayonda")
    await call.answer()

# ==========================
# STATUS: BAJARILDI
# ==========================
@dp.callback_query(F.data.startswith("done_"))
async def to_done(call: CallbackQuery):

    zid = int(call.data.split("_")[1])

    df = pd.read_excel(ZAYAVKA_FILE)
    df.loc[df["ID"] == zid, "Status"] = "Bajarildi"
    df.to_excel(ZAYAVKA_FILE, index=False)

    user_id = int(df[df["ID"] == zid]["UserID"].values[0])
    try:
        await bot.send_message(
            user_id,
            f"✅ Zayavkangiz #{zid} bajarildi."
        )
    except:
        pass

    await call.message.answer(f"✅ Zayavka #{zid} – Bajarildi")
    await call.answer()

# ==========================
# HISOBOT
# ==========================
@dp.message(Command("hisobot"))
async def send_excel(message: Message):
    if message.from_user.id not in ADMINS:
        return
    await message.answer_document(FSInputFile(ZAYAVKA_FILE))

# ==========================
# RUN
# ==========================
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
