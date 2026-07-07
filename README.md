# Etkinlik Personel Rapor Botu

Telegram uzerinden Excel dosyasi yukleyip personel bazli uye ve yatirim adetlerini raporlayan bot.

## Komutlar

- `/start` - Hos geldin mesaji
- `/yukle` - Excel yukleme ve rapor alma
- `/iptal` - Aktif islemi iptal etme

## Yerel calistirma

```bash
pip install -r requirements.txt
cp .env.example .env
# .env dosyasina TELEGRAM_BOT_TOKEN degerini yazin
python bot.py
```

Windows icin: `start_bot.bat` veya `python run_bot.py`

## Railway deploy

1. Bu repoyu Railway'e baglayin
2. **Variables** bolumune ekleyin:
   - `TELEGRAM_BOT_TOKEN` = BotFather token
3. Deploy baslatilir; `python bot.py` komutu ile worker olarak calisir
4. Ayni token ile baska yerde bot calistirmayin (409 Conflict olur)

## Excel mantigi

- Her personel ayri sheet'te
- **Uye Adedi:** `KAYIT TARİHİ` (E) secilen tarih araliginda
- **Yatirim Adedi:** `İLK YATIRIM TARİHİ` (F), kayit tarihinin ertesi gunu saat 10:00'a kadar

Tarih ornegi: `02.07.2026,04.07.2026` (baslangic ve bitis dahil)