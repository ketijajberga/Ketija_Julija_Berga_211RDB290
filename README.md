# VPN Risinājumu Mērījumu Dati

Šajā repozitorijā ir apkopoti VPN risinājumu (IPsec, Netmaker, OpenVPN, WireGuard, ZeroTier) testēšanas mērījumu dati.

## Mapju struktūra

| Mape | Saturs |
|------|--------|
| `1_Normali_tikla_apstakli_apstradati` | Normālu tīkla apstākļu statistikas kopsavilkums (Divu tīklu savienošana, Attālinātā piekļuve, Mākoņintegrācija) |
| `2_Normali_tikla_apstakli_neapstradati` | Normālu tīkla apstākļu pilnie mērījumi pa protokoliem |
| `3_Pasliktinati_tikla_apstakli_apstradati` | Pasliktinātu tīkla apstākļu statistikas kopsavilkums (Divu tīklu savienošana, Attālinātā piekļuve) |
| `4_Pasliktinati_tikla_apstakli_neapstradati` | Pasliktinātu tīkla apstākļu pilnie mērījumi pa protokoliem |
| `5_Kvalitativas_novertejumi` | Kvalitatīvā novērtēšana (konfigurācija, pārvaldāmība, drošība, izmaksas) |
| `6_SAW_rezultati` | SAW metodes aprēķini un galīgais novērtējums pa scenārijiem |

## Skripti

| Skripts | Apraksts |
|---------|----------|
| `VPN__normali_tikla_apstakli_statistika.py` | Aprēķina statistikas rādītājus (vidējā, mediāna, standartnovirze, min, maks, P95) no normālo apstākļu neapstrādātajiem mērījumiem un saglabā rezultātus Excel failā. |
| `VPN__pasliktinati_tikla_apstakli_statistika.py` | Aprēķina statistikas rādītājus no pasliktināto tīkla apstākļu neapstrādātajiem mērījumiem un saglabā rezultātus Excel failā. |
