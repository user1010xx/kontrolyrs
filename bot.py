import asyncio
import logging
import os
import shutil
import sys
import tempfile
import time
from io import BytesIO
from pathlib import Path

from dotenv import load_dotenv
from telegram import InputFile, Update
from telegram.error import Conflict, NetworkError, TimedOut
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    TypeHandler,
    filters,
)
from telegram.request import HTTPXRequest

from excel_processor import ProcessResult, parse_date_input, process_excel

load_dotenv()

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# Token'ı URL içinde basmasın; getUpdates gürültüsünü kes
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)
logging.getLogger("telegram.ext.Application").setLevel(logging.INFO)

WAITING_FILE, WAITING_DATES = range(2)
TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")

# Optional: comma-separated Telegram user IDs. Empty = allow everyone.
_ALLOWED_RAW = os.getenv("ALLOWED_USER_IDS", "").strip()
ALLOWED_USER_IDS: set[int] = {
    int(x.strip()) for x in _ALLOWED_RAW.split(",") if x.strip().isdigit()
}

# Limits
MAX_FILE_BYTES = int(os.getenv("MAX_FILE_MB", "20")) * 1024 * 1024
MAX_ROWS_PER_SHEET = int(os.getenv("MAX_ROWS_PER_SHEET", "100000"))
CONVERSATION_TIMEOUT_SEC = int(os.getenv("CONVERSATION_TIMEOUT_SEC", "900"))  # 15 dk
ALLOWED_EXTENSIONS = (".xlsx", ".xlsm")

# Telegram HTTP timeouts (default ~5s causes "Timed out" on big Excel / slow hosts)
HTTP_CONNECT_TIMEOUT = float(os.getenv("HTTP_CONNECT_TIMEOUT", "30"))
HTTP_READ_TIMEOUT = float(os.getenv("HTTP_READ_TIMEOUT", "120"))
HTTP_WRITE_TIMEOUT = float(os.getenv("HTTP_WRITE_TIMEOUT", "120"))
HTTP_POOL_TIMEOUT = float(os.getenv("HTTP_POOL_TIMEOUT", "30"))
SEND_RETRIES = int(os.getenv("SEND_RETRIES", "3"))


def _is_allowed(update: Update) -> bool:
    if not ALLOWED_USER_IDS:
        return True
    user = update.effective_user
    return bool(user and user.id in ALLOWED_USER_IDS)


async def _deny(update: Update) -> None:
    if update.effective_message:
        await update.effective_message.reply_text(
            "Bu botu kullanma yetkiniz yok. Yöneticinizle iletişime geçin."
        )


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not _is_allowed(update):
        await _deny(update)
        return ConversationHandler.END

    _cleanup(context)
    await update.message.reply_text(
        "Merhaba! Personel satış raporu botuna hoş geldiniz.\n\n"
        "Komutlar:\n"
        "/yukle — Excel yükle ve rapor al\n"
        "/iptal — Aktif işlemi iptal et\n"
        "/help — Yardım"
    )
    return ConversationHandler.END


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not _is_allowed(update):
        await _deny(update)
        return ConversationHandler.END
    await update.message.reply_text(
        "Kullanım:\n"
        "1) /yukle\n"
        "2) Excel dosyasını (.xlsx / .xlsm) doküman olarak gönderin\n"
        "3) Tarih aralığını yazın: 02.07.2026,04.07.2026\n\n"
        "Üye: KAYIT TARİHİ seçilen aralıkta\n"
        "Yatırım: İLK YATIRIM, kayıt anı ≤ yatırım ≤ ertesi gün 10:00\n\n"
        "/iptal ile işlemi iptal edebilirsiniz."
    )
    # Conversation içinde state koru
    if context.user_data.get("excel_path"):
        return WAITING_DATES
    if context.user_data.get("temp_dir") or context.user_data.get("_awaiting_file"):
        return WAITING_FILE
    return ConversationHandler.END


async def yukle(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not _is_allowed(update):
        await _deny(update)
        return ConversationHandler.END

    # Önceki temp dosyaları sil (clear() sızıntı yapmasın)
    _cleanup(context)
    context.user_data["_awaiting_file"] = True
    await update.message.reply_text(
        "Lütfen Excel dosyasını gönderin (.xlsx veya .xlsm, doküman olarak)."
    )
    return WAITING_FILE


async def receive_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not _is_allowed(update):
        await _deny(update)
        return ConversationHandler.END

    document = update.message.document
    if not document:
        await update.message.reply_text(
            "Lütfen bir Excel dosyasını doküman olarak gönderin (.xlsx / .xlsm)."
        )
        return WAITING_FILE

    raw_name = document.file_name or "dosya.xlsx"
    filename = Path(raw_name).name  # path traversal engeli
    lower = filename.lower()
    if not lower.endswith(ALLOWED_EXTENSIONS):
        await update.message.reply_text(
            "Sadece Excel dosyası kabul edilir: .xlsx veya .xlsm"
        )
        return WAITING_FILE

    file_size = document.file_size or 0
    if file_size > MAX_FILE_BYTES:
        mb = MAX_FILE_BYTES // (1024 * 1024)
        await update.message.reply_text(
            f"Dosya çok büyük (max {mb} MB). Daha küçük bir dosya gönderin."
        )
        return WAITING_FILE

    # Önceki dosyayı temizle, yenisini kaydet
    _cleanup(context)

    temp_dir = Path(tempfile.mkdtemp(prefix="etkinlik_"))
    file_path = temp_dir / filename
    try:
        tg_file = await document.get_file()
        await tg_file.download_to_drive(custom_path=str(file_path))
    except Exception:
        logger.exception("Dosya indirme hatasi")
        _cleanup_dir(temp_dir)
        await update.message.reply_text("Dosya indirilemedi. Lütfen tekrar deneyin.")
        return WAITING_FILE

    context.user_data["excel_path"] = str(file_path)
    context.user_data["temp_dir"] = str(temp_dir)
    context.user_data.pop("_awaiting_file", None)

    await update.message.reply_text(
        "Dosya alındı.\n\n"
        "Kontrol edilecek tarih aralığını yazın.\n"
        "Örnek: 02.07.2026,04.07.2026\n"
        "(Başlangıç ve bitiş tarihleri dahil)\n\n"
        "İptal için /iptal"
    )
    return WAITING_DATES


async def wrong_input_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(
        "Lütfen Excel dosyasını (.xlsx / .xlsm) Telegram'da *dosya/doküman* olarak gönderin.\n"
        "Fotoğraf veya metin kabul edilmez.",
        parse_mode=None,
    )
    return WAITING_FILE


async def wrong_input_dates(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text(
        "Şu an tarih aralığı bekleniyor.\n"
        "Örnek: 02.07.2026,04.07.2026\n"
        "İptal: /iptal — Yeni dosya: /yukle"
    )
    return WAITING_DATES


async def receive_dates(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    if not _is_allowed(update):
        await _deny(update)
        _cleanup(context)
        return ConversationHandler.END

    excel_path = context.user_data.get("excel_path")
    if not excel_path or not Path(excel_path).is_file():
        _cleanup(context)
        await update.message.reply_text(
            "Oturum süresi doldu veya dosya bulunamadı. Lütfen /yukle ile tekrar başlayın."
        )
        return ConversationHandler.END

    try:
        start_d, end_d = parse_date_input(update.message.text or "")
    except ValueError as exc:
        await update.message.reply_text(f"Hatalı tarih: {exc}")
        return WAITING_DATES

    await update.message.reply_text("Dosya işleniyor, lütfen bekleyin...")
    t0 = time.perf_counter()

    try:
        loop = asyncio.get_running_loop()
        result: ProcessResult = await loop.run_in_executor(
            None,
            lambda: process_excel(
                excel_path,
                start_d,
                end_d,
                include_zero_rows=True,
                max_rows_per_sheet=MAX_ROWS_PER_SHEET,
            ),
        )
        elapsed = time.perf_counter() - t0
        logger.info(
            "Excel islendi: %.1fs, personel=%s, uye=%s, yat=%s",
            elapsed,
            result.personel_count,
            result.total_uye,
            result.total_yat,
        )
        out_name = (
            f"rapor_{start_d.strftime('%d.%m.%Y')}_{end_d.strftime('%d.%m.%Y')}.xlsx"
        )
        caption_lines = [
            f"Tarih aralığı: {start_d.strftime('%d.%m.%Y')} - {end_d.strftime('%d.%m.%Y')} (dahil)",
            f"Toplam: {result.total_uye} üye, {result.total_yat} yatırım",
            f"Personel: {result.personel_count}",
        ]
        if result.warnings:
            # Telegram caption max ~1024
            warn_text = " | ".join(result.warnings)
            if len(warn_text) > 400:
                warn_text = warn_text[:397] + "..."
            caption_lines.append(f"Uyarı: {warn_text}")
        caption = "\n".join(caption_lines)
        if len(caption) > 1024:
            caption = caption[:1021] + "..."

        # BytesIO'yu kopyala; her denemede baştan okunabilsin
        payload = result.output.getvalue()
        await _send_document_with_retry(
            update,
            payload=payload,
            filename=out_name,
            caption=caption,
        )
        # Uzun uyarıları ayrı mesajda da gönder
        for w in result.warnings:
            if len(w) > 50:
                try:
                    await update.message.reply_text(
                        f"⚠️ {w[:3500]}",
                        read_timeout=HTTP_READ_TIMEOUT,
                        write_timeout=HTTP_WRITE_TIMEOUT,
                        connect_timeout=HTTP_CONNECT_TIMEOUT,
                    )
                except (TimedOut, NetworkError):
                    logger.warning("Uyari mesaji gonderilemedi: %s", w[:80])
    except (TimedOut, NetworkError) as exc:
        logger.exception("Telegram ag/timeout hatasi")
        try:
            await update.message.reply_text(
                "İşlem tamamlandı ancak Telegram bağlantısı zaman aşımına uğradı.\n"
                "Lütfen /yukle ile tekrar deneyin. Dosya çok büyükse biraz bekleyip yeniden gönderin.\n"
                f"Detay: {type(exc).__name__}",
                read_timeout=HTTP_READ_TIMEOUT,
                write_timeout=HTTP_WRITE_TIMEOUT,
                connect_timeout=HTTP_CONNECT_TIMEOUT,
            )
        except Exception:
            pass
    except Exception as exc:
        logger.exception("Excel işleme hatası")
        detail = str(exc).strip() or type(exc).__name__
        if len(detail) > 300:
            detail = detail[:297] + "..."
        try:
            await update.message.reply_text(
                "Dosya işlenirken hata oluştu.\n"
                f"Detay: {detail}\n\n"
                "Kontrol edin: sheet'lerde 'KAYIT TARİHİ' ve 'İLK YATIRIM TARİHİ' sütunları, "
                "tarih formatı (GG.AA.YYYY).",
                read_timeout=HTTP_READ_TIMEOUT,
                write_timeout=HTTP_WRITE_TIMEOUT,
                connect_timeout=HTTP_CONNECT_TIMEOUT,
            )
        except Exception:
            pass
    finally:
        _cleanup(context)

    return ConversationHandler.END


async def _send_document_with_retry(
    update: Update,
    *,
    payload: bytes,
    filename: str,
    caption: str,
) -> None:
    last_exc: Exception | None = None
    for attempt in range(1, SEND_RETRIES + 1):
        try:
            await update.message.reply_document(
                document=InputFile(BytesIO(payload), filename=filename),
                caption=caption,
                read_timeout=HTTP_READ_TIMEOUT,
                write_timeout=HTTP_WRITE_TIMEOUT,
                connect_timeout=HTTP_CONNECT_TIMEOUT,
                pool_timeout=HTTP_POOL_TIMEOUT,
            )
            return
        except (TimedOut, NetworkError) as exc:
            last_exc = exc
            logger.warning(
                "Rapor gonderimi deneme %s/%s basarisiz: %s",
                attempt,
                SEND_RETRIES,
                exc,
            )
            if attempt < SEND_RETRIES:
                await asyncio.sleep(1.5 * attempt)
    assert last_exc is not None
    raise last_exc


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    _cleanup(context)
    if update.effective_message:
        await update.effective_message.reply_text("İşlem iptal edildi.")
    return ConversationHandler.END


async def conversation_timeout(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    _cleanup(context)
    msg = update.effective_message if update else None
    if msg:
        await msg.reply_text(
            "Oturum zaman aşımına uğradı (dosya/tarih beklenirken işlem yok).\n"
            "Tekrar başlamak için /yukle"
        )
    return ConversationHandler.END


async def on_error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    if isinstance(context.error, Conflict):
        logger.error(
            "409 Conflict: ayni TELEGRAM_BOT_TOKEN ile baska bir bot calisiyor "
            "(yerel PC, ikinci Railway servisi, eski deploy). "
            "Digerini kapatin. Process 1 sn sonra cikiyor."
        )
        try:
            await context.application.stop()
        except Exception:
            logger.exception("Conflict sonrası stop başarısız")
        # Railway'in restart etmesi / logda net gorunmesi icin
        await asyncio.sleep(1)
        os._exit(1)
    logger.exception("Beklenmeyen hata", exc_info=context.error)
    if isinstance(update, Update) and update.effective_message:
        try:
            await update.effective_message.reply_text(
                "Beklenmeyen bir hata oluştu. /yukle ile tekrar deneyin."
            )
        except Exception:
            pass


async def log_incoming(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Her update'i logla — Railway'de 'mesaj geliyor mu?' teshisi icin."""
    user = update.effective_user
    chat = update.effective_chat
    msg = update.effective_message
    text = ""
    if msg is not None:
        text = (msg.text or msg.caption or "")[:120]
        if msg.document:
            text = f"[document:{msg.document.file_name}] {text}"
    logger.info(
        "Update: chat=%s user=%s (@%s) text=%r",
        chat.id if chat else None,
        user.id if user else None,
        getattr(user, "username", None),
        text,
    )


async def post_init(application: Application) -> None:
    me = await application.bot.get_me()
    logger.info("Bot oturum acildi: @%s id=%s", me.username, me.id)

    info = await application.bot.get_webhook_info()
    if info.url:
        logger.warning(
            "Webhook tanimli (%s) — long polling icin siliniyor.",
            info.url,
        )
        await application.bot.delete_webhook(drop_pending_updates=False)
    else:
        logger.info("Webhook yok; long polling kullanilacak.")

    if ALLOWED_USER_IDS:
        logger.info("ALLOWED_USER_IDS aktif: %s kullanici", len(ALLOWED_USER_IDS))
    else:
        logger.warning("ALLOWED_USER_IDS bos — bot herkese acik.")


def _cleanup_dir(temp_dir: str | Path | None) -> None:
    if not temp_dir:
        return
    path = Path(temp_dir)
    if not path.exists():
        return
    try:
        shutil.rmtree(path, ignore_errors=True)
    except Exception:
        logger.exception("Temp dizin silinemedi: %s", path)


def _cleanup(context: ContextTypes.DEFAULT_TYPE) -> None:
    temp_dir = context.user_data.pop("temp_dir", None)
    context.user_data.pop("excel_path", None)
    context.user_data.pop("_awaiting_file", None)
    _cleanup_dir(temp_dir)


def main() -> None:
    if not TOKEN:
        raise SystemExit("TELEGRAM_BOT_TOKEN bulunamadı. .env dosyasını kontrol edin.")

    token_preview = f"...{TOKEN[-6:]}" if len(TOKEN) > 6 else "(kisa)"
    logger.info("TELEGRAM_BOT_TOKEN yuklendi (%s)", token_preview)

    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

    request = HTTPXRequest(
        connection_pool_size=8,
        connect_timeout=HTTP_CONNECT_TIMEOUT,
        read_timeout=HTTP_READ_TIMEOUT,
        write_timeout=HTTP_WRITE_TIMEOUT,
        pool_timeout=HTTP_POOL_TIMEOUT,
    )
    # Long-poll getUpdates: read timeout polling suresinden uzun olmali
    get_updates_request = HTTPXRequest(
        connection_pool_size=4,
        connect_timeout=HTTP_CONNECT_TIMEOUT,
        read_timeout=max(HTTP_READ_TIMEOUT, 75.0),
        write_timeout=HTTP_WRITE_TIMEOUT,
        pool_timeout=HTTP_POOL_TIMEOUT,
    )
    app = (
        Application.builder()
        .token(TOKEN)
        .request(request)
        .get_updates_request(get_updates_request)
        .post_init(post_init)
        .build()
    )

    # conversation_timeout JobQueue ister; yoksa timeout'u kapat (entry/handlers calissin)
    use_timeout = CONVERSATION_TIMEOUT_SEC if app.job_queue is not None else None
    if use_timeout is None and CONVERSATION_TIMEOUT_SEC:
        logger.warning(
            "JobQueue yok — conversation_timeout devre disi. "
            "requirements: python-telegram-bot[job-queue]"
        )

    states: dict = {
        WAITING_FILE: [
            MessageHandler(filters.Document.ALL, receive_file),
            MessageHandler(~filters.COMMAND, wrong_input_file),
        ],
        WAITING_DATES: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, receive_dates),
            MessageHandler(~filters.COMMAND, wrong_input_dates),
        ],
    }
    if use_timeout is not None:
        states[ConversationHandler.TIMEOUT] = [
            MessageHandler(filters.ALL, conversation_timeout),
            CommandHandler("iptal", cancel),
        ]

    conv = ConversationHandler(
        entry_points=[CommandHandler("yukle", yukle)],
        states=states,
        fallbacks=[
            CommandHandler("iptal", cancel),
            CommandHandler("yukle", yukle),
            CommandHandler("start", start),
            CommandHandler("help", help_cmd),
        ],
        conversation_timeout=use_timeout,
        allow_reentry=True,
    )

    # group=-1: once logla (yanit vermese bile mesaj geldi mi gorunur)
    app.add_handler(TypeHandler(Update, log_incoming), group=-1)
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(conv)
    app.add_error_handler(on_error)

    logger.info(
        "Bot baslatiliyor (conv_timeout=%s, max_file=%sMB)...",
        use_timeout,
        MAX_FILE_BYTES // (1024 * 1024),
    )
    # drop_pending_updates=False: deploy sirasinda gonderilen /start kaybolmasin
    app.run_polling(
        allowed_updates=Update.ALL_TYPES,
        drop_pending_updates=False,
        bootstrap_retries=5,
    )


if __name__ == "__main__":
    main()
