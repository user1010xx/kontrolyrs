import asyncio
import logging
import os
import sys
import tempfile
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from telegram import Update
from telegram.error import Conflict
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)

from excel_processor import parse_date_input, process_excel

load_dotenv()

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

WAITING_FILE, WAITING_DATES = range(2)
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(
        "Merhaba! Personel satış raporu botuna hoş geldiniz.\n\n"
        "Rapor almak için /yukle komutunu kullanın."
    )


async def yukle(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.clear()
    await update.message.reply_text("Lütfen Excel dosyasını gönderin.")
    return WAITING_FILE


async def receive_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    document = update.message.document
    if not document:
        await update.message.reply_text("Lütfen bir Excel dosyası gönderin (.xlsx).")
        return WAITING_FILE

    filename = document.file_name or ""
    if not filename.lower().endswith((".xlsx", ".xlsm")):
        await update.message.reply_text("Sadece Excel dosyası (.xlsx) kabul edilir.")
        return WAITING_FILE

    temp_dir = Path(tempfile.mkdtemp(prefix="etkinlik_"))
    file_path = temp_dir / filename
    try:
        tg_file = await document.get_file()
        await tg_file.download_to_drive(custom_path=str(file_path))
    except Exception:
        logger.exception("Dosya indirme hatasi")
        _cleanup_dir(temp_dir)
        await update.message.reply_text("Dosya indirilemedi. Lutfen tekrar deneyin.")
        return WAITING_FILE

    context.user_data["excel_path"] = str(file_path)
    context.user_data["temp_dir"] = str(temp_dir)

    await update.message.reply_text(
        "Dosya alındı.\n\n"
        "Kontrol edilecek tarih aralığını yazın.\n"
        "Örnek: 02.07.2026,04.07.2026\n"
        "(Başlangıç ve bitiş tarihleri dahil)"
    )
    return WAITING_DATES


async def wrong_input_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Lutfen Excel dosyasini (.xlsx) dokuman olarak gonderin.")
    return WAITING_FILE


async def receive_dates(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    excel_path = context.user_data.get("excel_path")
    if not excel_path:
        await update.message.reply_text("Oturum süresi doldu. Lütfen /yukle ile tekrar başlayın.")
        return ConversationHandler.END

    try:
        start, end = parse_date_input(update.message.text)
    except ValueError as exc:
        await update.message.reply_text(f"Hatalı tarih: {exc}")
        return WAITING_DATES

    await update.message.reply_text("Dosya işleniyor, lütfen bekleyin...")

    try:
        output = process_excel(excel_path, start, end)
        out_name = f"rapor_{start.strftime('%d.%m.%Y')}_{end.strftime('%d.%m.%Y')}.xlsx"
        summary = pd.read_excel(output)
        output.seek(0)
        total_uye = int(summary["Üye Adedi"].sum())
        total_yat = int(summary["Yatırım Adedi"].sum())
        caption = (
            f"Tarih araligi: {start.strftime('%d.%m.%Y')} - {end.strftime('%d.%m.%Y')} (dahil)\n"
            f"Toplam: {total_uye} uye, {total_yat} yatirim"
        )
        await update.message.reply_document(
            document=output,
            filename=out_name,
            caption=caption,
        )
    except Exception:
        logger.exception("Excel işleme hatası")
        await update.message.reply_text("Dosya işlenirken hata oluştu. Dosya formatını kontrol edin.")
    finally:
        _cleanup(context)

    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    _cleanup(context)
    await update.message.reply_text("İşlem iptal edildi.")
    return ConversationHandler.END


async def on_error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    if isinstance(context.error, Conflict):
        logger.error(
            "Baska bir bot ornegi calisiyor. Once onu durdurun, sonra tekrar baslatin."
        )
        await context.application.stop()
        return
    logger.exception("Beklenmeyen hata", exc_info=context.error)


def _cleanup_dir(temp_dir: str | Path | None) -> None:
    if not temp_dir:
        return
    for path in Path(temp_dir).glob("*"):
        path.unlink(missing_ok=True)
    Path(temp_dir).rmdir()


def _cleanup(context: ContextTypes.DEFAULT_TYPE) -> None:
    temp_dir = context.user_data.pop("temp_dir", None)
    context.user_data.pop("excel_path", None)
    _cleanup_dir(temp_dir)


def main() -> None:
    if not TOKEN:
        raise SystemExit("TELEGRAM_BOT_TOKEN bulunamadı. .env dosyasını kontrol edin.")

    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.set_event_loop(asyncio.new_event_loop())

    app = Application.builder().token(TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("yukle", yukle)],
        states={
            WAITING_FILE: [
                MessageHandler(filters.Document.ALL, receive_file),
                MessageHandler(filters.TEXT & ~filters.COMMAND, wrong_input_file),
            ],
            WAITING_DATES: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_dates)],
        },
        fallbacks=[CommandHandler("iptal", cancel), CommandHandler("yukle", yukle)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv)
    app.add_error_handler(on_error)

    logger.info("Bot başlatılıyor...")
    app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)


if __name__ == "__main__":
    main()