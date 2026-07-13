# Etkinlik Personel Rapor Botu

Telegram uzerinden Excel dosyasi yukleyip personel bazli uye ve yatirim adetlerini raporlayan bot.

## Komutlar

- `/start` — Hos geldin / aktif islemi sifirlar
- `/yukle` — Excel yukleme ve rapor alma
- `/iptal` — Aktif islemi iptal etme
- `/help` — Kullanim ozeti

## Yerel calistirma

```bash
pip install -r requirements.txt
# Windows: copy .env.example .env
# Linux/macOS: cp .env.example .env
# .env dosyasina TELEGRAM_BOT_TOKEN degerini yazin
python bot.py
```

Windows icin: `start_bot.bat`, `start_bot.ps1` veya `python run_bot.py`

### Ortam degiskenleri

| Degisken | Aciklama |
|----------|----------|
| `TELEGRAM_BOT_TOKEN` | Zorunlu — BotFather token |
| `ALLOWED_USER_IDS` | Opsiyonel — virgulle user ID; bos = herkese acik |
| `MAX_FILE_MB` | Varsayilan 20 |
| `MAX_ROWS_PER_SHEET` | Varsayilan 100000 |
| `CONVERSATION_TIMEOUT_SEC` | Varsayilan 900 (15 dk) |

## Railway deploy

1. Bu repoyu Railway'e baglayin
2. **Variables** bolumune ekleyin:
   - `TELEGRAM_BOT_TOKEN` = BotFather token
   - (onerilir) `ALLOWED_USER_IDS` = yetkili Telegram ID'ler
3. Deploy baslatilir; `python bot.py` worker olarak calisir
4. **Tek replica** kullanin; ayni token ile baska yerde bot calistirmayin (409 Conflict)
5. Yerel `python bot.py` / `run_bot.py` calisiyorsa **kapatin** — aksi halde Railway mesaj alamaz
6. Deploy loglarinda sunlari arayin:
   - `Bot oturum acildi: @yarismakt_bot` (dogru bot mu?)
   - `Update: chat=... text='/start'` (mesaj geliyor mu?)
   - `409 Conflict` varsa baska instance token'i kullaniyor

## Excel mantigi

- Her personel ayri sheet'te
- `TOPLAM` / `MANUEL EKLENENLER` (buyuk/kucuk harf duyarsiz) atlanir
- **Uye Adedi:** `KAYIT TARİHİ` (veya benzer baslik) secilen tarih araliginda
- **Yatirim Adedi:** `İLK YATIRIM TARİHİ`, kayit anindan itibaren **ertesi gun 10:00'a kadar** (dahil); kayittan onceki yatirim sayilmaz

Tarih ornegi: `02.07.2026,04.07.2026` (baslangic ve bitis dahil)

Desteklenen tarih ornekleri: `04.07.2026 10:57`, saniyeli, ISO (`2026-07-04T10:57:00`), Excel seri numarasi, bos hucre (NaT guvenli)

## Test

```bash
python -m unittest discover -s tests -v
```
