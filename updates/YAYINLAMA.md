# İlk yayın adımları

1. Yalnızca bu klasörün içeriğini yeni GitHub deposuna ekleyin. Kişisel `outputs\Dosya` verilerini eklemeyin.
2. Depo **public** olacaksa kaynak kodu ve Release dosyaları herkes tarafından görülebilir. Depo **private** kalırsa uygulamanın anonim güncelleme indirmesi çalışmaz; ayrı bir herkese açık indirme noktası gerekir.
3. Depoda **Settings > Pages > Build and deployment > Source** değerini **GitHub Actions** yapın.
4. Kaynak değişikliğini varsayılan dala gönderin.
5. `v1.0.1` gibi bir etiket oluşturup gönderin. `.github/workflows/release.yml` Windows EXE'lerini üretir, Release'e ekler ve güncelleme sayfasını yayımlar.
6. Uygulamayı ilk defa dağıtırken `DosyaGezgini.exe` ile `DosyaGezgini-Updater.exe` dosyalarını aynı klasöre koyun. Sonraki sürümlerde uygulama güncelleme denetimini kullanabilir.

Yeni bilgisayardaki uygulama sürüm kontrolünü Ayarlar bölümünde, depoya göre derlenmiş güncelleme adresi üzerinden yapar.
