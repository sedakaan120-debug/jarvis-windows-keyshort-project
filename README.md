# J.A.R.V.I.S. Voice Assistant & Keyshort Engine

Bu proje, orijinal olarak **Alp Ünlü** tarafından geliştirilen J.A.R.V.I.S. Sesli Asistan projesinin **Windows işletim sistemine özel olarak port edilmiş** ve üzerine gelişmiş otomasyon katmanları eklenmiş güncel sürümüdür. Ana asistan iskeletini topluluğa sunarak bu geliştirmeye zemin hazırlayan **Alp Ünlü**'ye teşekkür ve credit borç biliriz.

---

## 🚀 Kaan tarafından geliştirilen Windows Uzantısı: Keyshort Makro Entegrasyonu

Bu sürüm, Windows API ve sistem mimarisiyle tam uyumlu çalışacak şekilde **Kaan** tarafından tamamen sıfırdan geliştirilip projeye enjekte edilmiş dinamik bir **Keyshort Makro Motoru** ve **Arayüz Sekmesi** içerir. 

Bu Windows entegrasyonu sayesinde asistan, işletim sistemi genelinde (`global hotkey`) klavye kısayollarını arka planda dinler. Atanan tuş kombinasyonlarına basıldığında Windows üzerindeki herhangi bir uygulamayı (`.exe`), masaüstü kısayolunu (`.lnk`), toplu iş dosyasını (`.bat`) veya web URL'sini saliseler içinde tetikler.

### 🛠️ Öne Çıkan Özellikler
* **Gelişmiş Makro Sekmesi:** J.A.R.V.I.S. arayüzüne tamamen entegre, tema uyumlu özel yönetim paneli.
* **Windows Global Tuş Dinleme:** `keyboard` kütüphanesi yardımıyla, arka planda başka bir oyunda veya uygulamada olsanız dahi Windows genelinde çalışan tetikleyiciler.
* **Görsel Dosya Seçici:** Kullanıcıların çalıştırılacak Windows dosyalarını (`.exe`, `.lnk`, `.bat`) bilgisayardan kolayca seçebilmesi için `filedialog` entegrasyonu.
* **Anlık Hafıza Yönetimi:** Eklenen veya silinen makrolar anında `config/macros.json` dosyasına işlenir; arka plan motoru thread'i döngüsel olarak bu dosyayı tarar ve sistemi yeniden başlatmaya gerek kalmadan tuşları dinamik olarak hafızaya alır veya serbest bırakır.
* **Güvenli Silme Motoru:** Atanmış bir kısayolu arayüzden tek tuşla silme ve Windows işletim sistemi tuş dinleyicisinden (`unhotkey`) temizleme desteği.

### 📖 Kullanım Rehberi

1. **Makro Ekleme:**
   * Arayüzdeki **MACROS** sekmesine geçin.
   * *Tetikleyici Tuş* kutusuna kombinasyonu yazın (Örn: `ctrl+shift+v` veya `windows+alt+s`).
   * *Çalıştırılacak Hedef* kısmından **SEÇ** butonuna basarak tetiklemek istediğiniz uygulamayı veya kısayolu seçin.
   * **MAKROYU KAYDET** butonuna basın. Sol panelde onay logunu göreceksiniz.

2. **Makro Silme:**
   * **MACROS** sekmesindeki *Tetikleyici Tuş* kutusuna silmek istediğiniz kombinasyonu yazın (Örn: `ctrl+shift+v`).
   * Kırmızı renkli **MAKROYU SİL** butonuna basın. Windows işletim sistemi tuş dinleyicisi tuşu anında serbest bırakacaktır.

### 📂 Projeye Eklenen Katmanlar
* `actions/macros.py` - Klavye dinleme, thread yönetimi ve dinamik JSON okuma/yazma motoru.
* `ui.py` (Güncellenen Kısım) - Makro yönetim sekmesi, Tkinter buton ve giriş bileşenleri, silme/kaydetme arayüz fonksiyonları.
* `config/macros.json` - Makroların haritalanmış verilerini Windows üzerinde güvenli bir şekilde tutan yerel veritabanı dosyası.