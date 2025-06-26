from graphviz import Digraph

# Akış diyagramını oluştur
dot = Digraph("Fatura Otomasyonu", format="png")

# Düğümleri ekle
dot.node("A", "Fatura Alma (PDF/Resim)")
dot.node("B", "OCR ile Metin Tanıma")
dot.node("C", "Veri Sınıflandırma\n(Standart, Firma, Kullanıcı, AI)")
dot.node("D", "Zorunlu Alanları Kontrol Et\n(GTİP, Menşei, Adet vb.)")
dot.node("E", "Eksik Alanları Kullanıcıya Göster")
dot.node("F", "AI ile GTİP Tahmini ve Veri Doğrulama")
dot.node("G", "Beyanname Formatına Dönüştür")
dot.node("H", "Gümrük Sistemine Aktar")
dot.node("I", "İşlem Tamamlandı")

# Kenarları ekle
dot.edge("A", "B")
dot.edge("B", "C")
dot.edge("C", "D")
dot.edge("D", "E", label="Eksik Alan Var?")
dot.edge("E", "F", label="Evet")
dot.edge("E", "G", label="Hayır")
dot.edge("F", "G")
dot.edge("G", "H")
dot.edge("H", "I")

# Diyagramı oluştur
diagram_path = "/mnt/data/fatura_otomasyonu.png"
dot.render(diagram_path, format="png")

# Görseli göster
diagram_path
