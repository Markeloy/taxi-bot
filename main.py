import logging
import os
from datetime import datetime, date, timedelta
from aiogram import Bot, Dispatcher, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.utils import executor
from aiogram.dispatcher import FSMContext
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher.filters.state import State, StatesGroup
from dotenv import load_dotenv
import xlsxwriter

load_dotenv()

API_TOKEN = os.getenv('API_TOKEN')
ADMIN_ID = int(os.getenv('ADMIN_ID'))

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot, storage=MemoryStorage())
logging.basicConfig(level=logging.INFO)

class OrderTaxi(StatesGroup):
    location = State()
    last_name = State()
    hospital_target = State()
    time = State()
    passengers = State()
    destination = State()

class AwaitingAddress(StatesGroup):
    waiting_for_address = State()
    selecting_order = State()

orders = {}

driver_list = [
    {"id": 111111111, "car": "KIA RIO", "plate": "", "color": "Белая"},
    {"id": 222222222, "car": "INFINITI M35X", "plate": "", "color": "Серая"},
    {"id": 333333333, "car": "Volkswagen Tiguan", "plate": "", "color": "Белый"},
    {"id": 444444444, "car": "Geely", "plate": "", "color": "Серая"},
    {"id": 555555555, "car": "Mitsubishi Outlander", "plate": "", "color": "Оранжевая"}
]

start_kb = ReplyKeyboardMarkup(resize_keyboard=True)
start_kb.add(KeyboardButton("Заказать такси"))

location_kb = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
location_kb.add("РусАгроГрупп (Завод)", "ЦРБ (Больница)")
location_kb.add("Другой адрес", "Назад")

time_kb = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
time_kb.add("Сейчас", "Выбрать время позже", "Назад")

hospital_kb = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
hospital_kb.add("Улица Туркова 4", "Улица Туркова 10", "Назад")

@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    if message.from_user.id == ADMIN_ID:
        kb = ReplyKeyboardMarkup(resize_keyboard=True)
        kb.add(KeyboardButton("Заказать такси"))
        kb.add(KeyboardButton("📋 Заказы за месяц"))
    else:
        kb = start_kb
    await message.answer("Привет! Я помогу тебе вызвать такси.", reply_markup=kb)

@dp.message_handler(lambda message: message.text == "Заказать такси", state=None)
async def order_start(message: types.Message):
    await message.answer("Откуда вас забрать?", reply_markup=location_kb)
    await OrderTaxi.location.set()

@dp.message_handler(state=OrderTaxi.location)
async def choose_location(message: types.Message, state: FSMContext):
    if message.text in ["Назад", "Заказать такси"]:
        await state.finish()
        await start(message)
        return
    location = message.text
    await state.update_data(location=location)

    if location == "Другой адрес":
        await message.answer("Введите адрес вручную:")
        return
    if location == "РусАгроГрупп (Завод)":
        await message.answer("Введите вашу фамилию:")
        await OrderTaxi.last_name.set()
        return
    if location == "ЦРБ (Больница)":
        await message.answer("Выберите адрес прибытия:", reply_markup=hospital_kb)
        await OrderTaxi.hospital_target.set()
        return

    await message.answer("Когда подать машину?", reply_markup=time_kb)
    await OrderTaxi.time.set()

@dp.message_handler(state=OrderTaxi.last_name)
async def get_last_name(message: types.Message, state: FSMContext):
    if message.text == "Назад":
        await order_start(message)
        return
    await state.update_data(last_name=message.text, time="19:30")
    await message.answer("Сколько пассажиров?", reply_markup=ReplyKeyboardMarkup(resize_keyboard=True).add("Назад"))
    await OrderTaxi.passengers.set()

@dp.message_handler(state=OrderTaxi.hospital_target)
async def hospital_target_choice(message: types.Message, state: FSMContext):
    if message.text == "Назад":
        await order_start(message)
        return
    await state.update_data(destination=message.text)
    await message.answer("Когда подать машину?", reply_markup=time_kb)
    await OrderTaxi.time.set()

@dp.message_handler(state=OrderTaxi.time)
async def choose_time(message: types.Message, state: FSMContext):
    if message.text == "Назад":
        await order_start(message)
        return
    await state.update_data(time=message.text)
    await message.answer("Сколько пассажиров?", reply_markup=ReplyKeyboardMarkup(resize_keyboard=True).add("Назад"))
    await OrderTaxi.passengers.set()

@dp.message_handler(state=OrderTaxi.passengers)
async def choose_passengers(message: types.Message, state: FSMContext):
    if message.text == "Назад":
        await order_start(message)
        return
    passengers = message.text
    if not passengers.isdigit():
        await message.answer("Пожалуйста, введите количество пассажиров числом.")
        return
    passengers = int(passengers)
    data = await state.get_data()
    location = data.get('location', '-')
    time = data.get('time', '-')
    username = message.from_user.username or message.from_user.full_name

    cars_needed = 1
    if location == "РусАгроГрупп (Завод)":
        if passengers > 8:
            cars_needed = 3
        elif passengers > 4:
            cars_needed = 2

    order_id = message.from_user.id
    orders[order_id] = {
        'client': username,
        'location': location,
        'last_name': data.get('last_name', ''),
        'time': time,
        'passengers': passengers,
        'cars_needed': cars_needed,
        'destination': data.get('destination', ''),
        'status': 'pending',
        'created': datetime.now().date()
    }

    order_summary = (
        f"🚖 Новый заказ такси\n"
        f"Клиент: @{username} {data.get('last_name', '')}\n"
        f"Откуда: {location}\n"
        f"Время: {time}\n"
        f"Пассажиров: {passengers}\n"
        f"Нужно машин: {cars_needed}\n\n"
        f"Назначьте водителя: /assign_{order_id}"
    )
    await bot.send_message(ADMIN_ID, order_summary)
    await message.answer("Спасибо! Ожидайте, мы ищем машину.", reply_markup=start_kb)
    await state.finish()

@dp.message_handler(lambda message: message.text.startswith("/assign_"))
async def assign_driver(message: types.Message):
    if message.from_user.id != ADMIN_ID:
        return
    try:
        order_id = int(message.text.replace("/assign_", ""))
        order = orders.get(order_id)
        if not order:
            await message.answer("⚠️ Заказ не найден.")
            return
        keyboard = InlineKeyboardMarkup()
        for driver in driver_list:
            btn_label = f"{driver.get('car')} {driver.get('color', '')}".strip()
            keyboard.add(InlineKeyboardButton(
                text=btn_label,
                callback_data=f"assign|{order_id}|{driver['id']}"
            ))
        await message.answer("Выберите водителя для этого заказа:", reply_markup=keyboard)
    except Exception as e:
        await message.answer(f"Ошибка: {str(e)}")

@dp.callback_query_handler(lambda c: c.data.startswith("assign|"))
async def confirm_assignment(callback_query: types.CallbackQuery, state: FSMContext):
    _, order_id, driver_id = callback_query.data.split("|")
    order_id = int(order_id)
    driver_id = int(driver_id)
    order = orders.get(order_id)

    if not order:
        await callback_query.answer("⚠️ Заказ не найден.", show_alert=True)
        return

    driver = next((d for d in driver_list if d["id"] == driver_id), None)
    if not driver:
        await callback_query.answer("⚠️ Водитель не найден.", show_alert=True)
        return

    assigned = orders[order_id].setdefault("drivers", [])
    if driver_id in assigned:
        await callback_query.answer("Этот водитель уже назначен.", show_alert=True)
        return

    assigned.append(driver_id)
    orders[order_id]["status"] = "assigned"

    await bot.send_message(order_id,
        f"🚕 Ваш водитель: {driver.get('car')} {driver.get('color', '')}\nОжидайте на месте подачи!")
    await callback_query.answer("✅ Водитель добавлен.")

    total_needed = orders[order_id].get("cars_needed", 1)
    if len(assigned) < total_needed:
        await bot.send_message(ADMIN_ID,
            f"✅ Водитель {driver.get('car')} добавлен. Назначьте ещё {total_needed - len(assigned)} машин(ы) для заказа ID {order_id}.")
        return

    kb = ReplyKeyboardMarkup(resize_keyboard=True).add("📝 Ввести адреса клиентов")
    await bot.send_message(ADMIN_ID,
        f"🟢 Все водители назначены на заказ @{order['client']} (ID {order_id})",
        reply_markup=kb)
    await state.update_data(order_id=order_id)
    await AwaitingAddress.selecting_order.set()

@dp.message_handler(lambda m: m.text == "📝 Ввести адреса клиентов", state=AwaitingAddress.selecting_order)
async def prompt_manual_address_entry(message: types.Message, state: FSMContext):
    await message.answer("✏️ Введите адреса клиентов (в одном сообщении или по очереди):")
    await AwaitingAddress.waiting_for_address.set()

@dp.message_handler(lambda m: not m.text.startswith("/") and m.text != "✅ Завершить заказ", state=AwaitingAddress.waiting_for_address)
async def set_custom_address(message: types.Message, state: FSMContext):
    if message.from_user.id != ADMIN_ID:
        return
    data = await state.get_data()
    order_id = data.get("order_id")
    if not order_id or order_id not in orders:
        await message.answer("⚠️ Ошибка: заказ не найден.")
        await state.finish()
        return
    destination_list = orders[order_id].setdefault("destination_list", [])
    destination_list.append(message.text.strip())
    orders[order_id]["status"] = "completed"
    kb = ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add("✅ Завершить заказ")
    await message.answer(
        f"✅ Адрес добавлен: {message.text.strip()}\nВы можете добавить ещё адрес или нажмите '✅ Завершить заказ'.",
        reply_markup=kb
    )

@dp.message_handler(lambda m: m.text == "📋 Заказы за месяц")
async def list_orders_by_admin(message: types.Message):
    if message.from_user.id != ADMIN_ID:
        return

    today = date.today()
    month_ago = today - timedelta(days=30)
    filename = f"orders_report_{today.isoformat()}.xlsx"
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    headers = ["Дата", "Клиент", "Фамилия", "Откуда", "Куда", "Пассажиров", "Машин", "Статус", "Доп. адреса"]
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    row = 1
    to_delete = []
    for order_id, order in orders.items():
        created = order.get("created")
        if not created:
            continue
        if (today - created).days > 35:
            to_delete.append(order_id)
            continue
        if created >= month_ago:
            worksheet.write(row, 0, str(created))
            worksheet.write(row, 1, f"@{order['client']}")
            worksheet.write(row, 2, order.get('last_name', ''))
            worksheet.write(row, 3, order['location'])
            worksheet.write(row, 4, order.get('destination', '-'))
            worksheet.write(row, 5, order['passengers'])
            worksheet.write(row, 6, order['cars_needed'])
            worksheet.write(row, 7, order['status'])
            worksheet.write(row, 8, ", ".join(order.get("destination_list", [])))
            row += 1

    workbook.close()

    with open(filename, "rb") as f:
        await bot.send_document(message.chat.id, f)

    os.remove(filename)
    for oid in to_delete:
        del orders[oid]

@dp.message_handler(lambda m: m.text == "✅ Завершить заказ", state=AwaitingAddress.waiting_for_address)
async def finish_manual_input(message: types.Message, state: FSMContext):
    await message.answer("🟢 Заказ завершён.", reply_markup=start_kb)
    await state.finish()

if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
