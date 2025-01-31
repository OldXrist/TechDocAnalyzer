import logging
import sys
from os import getenv

import httpx
from dotenv import load_dotenv

from aiogram import Router, html
from aiogram.filters import CommandStart, Command
from aiogram.types import Message

from filters.admin import IsAdminFilter

load_dotenv()
GROQ_API_KEY = getenv("GROQ_API_KEY")


# Create a router for command handlers
router = Router()

load_dotenv()
ADMIN_IDS = list(map(int, getenv("ADMIN_IDS").split(',')))

logging.basicConfig(level=logging.INFO, stream=sys.stdout)
logger = logging.getLogger(__name__)


@router.message(CommandStart(), IsAdminFilter(ADMIN_IDS))
async def command_start_handler(message: Message) -> None:
    await message.answer(f"Добро пожаловать, {html.bold(message.from_user.first_name)}!"
                         f"\n\nЯ ваш личный ИИ-ассистент, чем я могу вам помочь?")


@router.message(CommandStart())
async def command_start_handler(message: Message) -> None:
    await message.answer(f"Здавствуйте, {html.bold(message.from_user.first_name)}!"
                         f"\n\nК сожалению, вы не авторизованы для использования этого бота."
                         f"\n\nЕсли вы считаете, что это ошибка, свяжитесь с администратором.")


@router.message(Command('rate_limit'), IsAdminFilter(ADMIN_IDS))
async def check_openai_rate_limit(message: Message):
    """Checks the remaining rate limit for Groq API."""
    try:
        # Send a test request to Groq API
        async with httpx.AsyncClient() as client:
            response = await client.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers={"Authorization": f"Bearer {GROQ_API_KEY}"},
                json={
                    "model": "llama-3.3-70b-versatile",
                    "messages": [{"role": "system", "content": "Test request"}],
                },
            )

        # Extract rate limit information from headers
        remaining_requests = response.headers.get("x-ratelimit-remaining-requests", "Unknown")
        limit_reset_time = response.headers.get("x-ratelimit-reset-requests", "Unknown")

        # Format response message
        reply_message = (
            f"📊 *Информация о лимите запросов:*\n"
            f"- Осталось запросов: `{remaining_requests}`\n"
            f"- Время сброса лимита: `{limit_reset_time}`\n"
            f"⏳ Пожалуйста, учитывайте это при использовании!!"
        )
    except Exception as e:
        # Handle any errors
        logger.error(f'Request error: {e}')
        reply_message = f"❌ Ошибка при отправке запроса."

    # Send the message to the user
    await message.answer(reply_message, parse_mode="Markdown")
