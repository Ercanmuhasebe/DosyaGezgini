# Dosya Gezgini bulut derlemesi

Bu klasör, uygulamanın kaynak kodunu, Windows üzerinde çalışan otomatik derleme akışını ve güncelleme denetimi sayfasını içerir. `outputs\Dosya` altındaki bilgisayara özel ayarlar, geçmiş, kısayollar ve yedekler bu pakete alınmamıştır.

## ChatGPT Codex Cloud ile çalışma

Codex Cloud, bağlı bir GitHub deposundaki kaynak kodla çalışır. Bu proje klasörü depoya eklendikten ve depo ChatGPT/Codex hesabında etkinleştirildikten sonra bulut görevleri bu kaynak üzerinde değişiklik yapabilir. EXE'yi dağıtmak ve uygulamanın güncelleme adresini sağlamak için GitHub Actions, Releases ve Pages kullanılır; ChatGPT sohbeti tek başına herkese açık EXE indirme adresi sağlamaz.

## Her sürümde otomatik derleme

Projeyi GitHub deposuna aldıktan ve depo ayarlarında **Settings > Pages > Build and deployment > Source: GitHub Actions** seçtikten sonra, `v1.0.1` biçiminde yeni bir sürüm etiketi yayımlandığında akış Windows derlemesini yapar. Uygulama ve güncelleme yardımcısı EXE'lerini, ilk kurulum ZIP dosyasını Release'e koyar; SHA-256 değerini hesaplar ve `latest.xml` dosyasını Pages üzerinde yayımlar.

Sürüm yayımlamak için GitHub'da kaynak değişikliklerini ana dala gönderin, ardından `v1.0.1`, `v1.0.2` gibi yeni bir etiket oluşturup gönderin. Uygulamanın sürüm numarası etiketten alınır; güncelleme adresi derleme sırasında depo adına göre oluşturulur.

## Yerel derleme

Windows'ta PowerShell ile:

```powershell
./build.ps1 -Version 1.0.1 -Repository Ercanmuhasebe/DosyaGezgini
```

Çıktılar `dist` klasöründe oluşur. İlk kurulumda iki EXE'yi mevcut program klasörüne kopyalayıp eskilerin üzerine yazın; eski `Dosya` klasörünü silmeyin. İnternet üzerinden otomatik güncelleme için uygulama ve `DosyaGezgini-Updater.exe` aynı klasörde bulunmalıdır.

## Yayın öncesi

İlk yayından önce GitHub deposunun herkese açık mı özel mi olacağına karar verin. Güncelleme dosyasının diğer bilgisayarlardan kimlik doğrulaması olmadan indirilebilmesi gerekir. Bu pakete bilgisayara özel `Dosya` verileri veya müşteri belgeleri dahil edilmemiştir.
