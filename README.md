# CIRIBU – VSTO dodatak za konverziju ćirilice u latinicu

**CIRIBU** je Microsoft Word dodatak razvijen kao VSTO (Visual Studio Tools for Office) projekat u C#. Omogućava korisnicima da brzo i lako konvertuju srpski ćirilični tekst u latinicu, direktno u dokumentu.

📎 *Naziv dodatka: CIRIBU*  
📎 *Verzija: 1.0.0.2*  
📎 *Target aplikacija: Microsoft Word (Office 2013–2019 / 365)*  
📎 *Platforma: AnyCPU (.NET 4.7.2, VSTO Runtime 4.0)*

---

## 🛠️ Funkcionalnosti

- ✅ Brza konverzija ćiriličnog teksta u srpsku latinicu
- ✅ Integracija sa Word interfejsom (Ribbon dugme)
- ✅ Nema eksternih zavisnosti osim Office Interop biblioteka
- ✅ Automatsko ažuriranje dodatka omogućeno

---

## 📦 Tehnologije

- C# (.NET Framework 4.7.2)
- Microsoft VSTO (Visual Studio Tools for Office)
- Office Interop API (Word)
- Visual Studio 2019/2022

---

## 🧭 Instalacija

1. Build-uj projekat u **Release** modu.
2. Pokreni `setup.exe` iz `publish` foldera (generisan tokom publish procesa).
3. Osiguraj da su instalirani:
   - .NET Framework 4.7.2
   - VSTO Runtime 4.0 (x86/x64)

Nakon instalacije, dodatak će se automatski pojaviti u Word-u.

---

## 🚀 Korišćenje

1. Otvori Microsoft Word.
2. Na Ribbon traci pojaviće se nova sekcija sa dugmetom **"Ćirilica → Latinica"**.
3. Selektuj ćirilični tekst u dokumentu i klikni dugme.
4. Tekst se automatski konvertuje u srpsku latinicu.

---

## 🔐 Licenca

Ovaj softver je besplatan za ličnu i nekomercijalnu upotrebu.

📄 Projekat je licenciran pod **MIT licencom** – pogledaj fajl [LICENSE](LICENSE) za detalje.

---

## 💼 Komercijalna upotreba

Ako planiraš da koristiš ovaj dodatak u okviru komercijalnog rešenja, molim te da me kontaktiraš radi dogovora:

📧 **sreckojovancevic@gmail.com**

---

## 👤 Autor

Ovaj dodatak razvio je **Srećko Jovančević**, iz potrebe da unapredi svakodnevni rad u Word-u jednostavnim i pouzdanim alatom za konverziju teksta.

Tokom rada korišćena je i pomoć veštačke inteligencije – **ChatGPT (OpenAI)** – u pisanju dokumentacije, pisanju koda i rešavanju tehničkih izazova.

> **Ovaj dodatak je plod ljudske volje i tehničke saradnje – uz poštovanje prema znanju, vremenu i zajednici.**

---

## 💬 Kontakt i podrška

Za sva pitanja, predloge, ili prijave grešaka – slobodno otvori **issue** na ovom repozitorijumu ili me kontaktiraj putem e-pošte.
