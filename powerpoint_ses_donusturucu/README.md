# PowerPoint'ten Sesli Anlatıma Dönüştürücü

**PowerPoint'ten Sesli Anlatıma Dönüştürücü**, PowerPoint sunumlarını (PPTX) alır ve istediğiniz dilde sesli anlatıma dönüştürür. Bu program, kullanıcıların PowerPoint sunumlarına sesli anlatımlar eklemelerini kolaylaştırarak sunumlarını daha etkili ve etkileşimli hale getirir. Arapça, İngilizce, Rusça ve diğer birçok dilde sesli anlatım desteği sunar.

## Sahibi ve Geliştiricisi

Bu proje, **Mehmet Yay** tarafından geliştirilmiştir. Mehmet Yay, bu yazılımın sahibi ve geliştiricisidir.

## Özellikler

- **PPTX Dosyasını Yükleyin**: Kullanıcılar bir PowerPoint dosyasını seçer ve yazılım onu alır.
- **Birden Çok Dil Seçeneği**: Sesli anlatımı Arapça, İngilizce, Rusça gibi çeşitli dillerde oluşturabilirsiniz.
- **Google TTS Entegresi**: **gTTS (Google Text-to-Speech)** kütüphanesi kullanılarak yüksek kaliteli sesli anlatım oluşturulur.
- **Kolay Kullanım**: Kullanıcı dostu arayüz ile birkaç tıklama ile PowerPoint dosyasını sesli anlatıma dönüştürün.

## Gereksinimler

Projenin çalışabilmesi için aşağıdaki Python kütüphanelerinin yüklenmiş olması gerekir:

- **Pillow**
- **pygame**
- **gTTS**
- **pydub**
- **requests**
- **beautifulsoup4**
- **psutil**
- **pyautogui**
- **python-pptx**
- **packaging**
- **zipfile36**

Tüm bu kütüphaneleri yüklemek için, proje klasöründe terminal veya komut satırında şu komutu çalıştırabilirsiniz:

```bash
pip install -r requirements.txt
