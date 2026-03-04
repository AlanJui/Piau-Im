import re

with open('docs/_template.html', 'r', encoding='utf-8') as f:
    html = f.read()

html = re.sub(r'<style id=\"injected-design\">.*?</style>', '', html, flags=re.DOTALL)

style_block = '''<style id=\"injected-design\">
/* =========================
   頁面全域邊距與大字體 (Index & Articles)
   ========================= */
body {
    padding: 0 50px !important;
}

@media (max-width: 768px) {
    body {
        padding: 0 20px !important;
    }
}

/* =========================
   Index頁面 Header & Nav 改良
   ========================= */
#header {
  text-align: center;
  padding: 60px 20px 40px !important;
  background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%) !important;
  color: white !important;
  border-radius: 0 0 20px 20px !important;
  margin-bottom: 40px !important;
  box-shadow: 0 8px 25px rgba(0,0,0,0.15) !important;
}

#header img {
  border-radius: 10px;
  box-shadow: 0 4px 10px rgba(0,0,0,0.2);
  margin-bottom: 20px;
  max-width: 100%;
  height: auto;
}

#header h1 {
  font-family: "Noto Serif TC", serif !important;
  font-weight: 900 !important; /* Noto Serif TC Black */
  font-size: 48pt !important;
  margin: 0 !important;
  text-shadow: 3px 3px 6px rgba(0,0,0,0.4) !important;
  letter-spacing: 5px !important;
}

.main-nav {
  display: flex !important;
  justify-content: center;
  align-items: center;
  margin-bottom: 40px;
  background: white;
  padding: 15px 30px;
  border-radius: 50px;
  box-shadow: 0 4px 15px rgba(0,0,0,0.05);
  max-width: 800px;
  margin-left: auto;
  margin-right: auto;
}

.main-nav ul {
  list-style: none;
  padding: 0;
  margin: 0;
  display: flex;
  gap: 30px;
}

.main-nav ul li a {
  text-decoration: none;
  color: #1e3c72 !important;
  font-family: "Noto Serif TC", serif !important;
  font-weight: 700 !important;
  font-size: 24pt !important;
  padding: 10px 24px;
  border-radius: 30px;
  transition: background 0.3s, color 0.3s;
}

.main-nav ul li a:hover {
  background: #1e3c72 !important;
  color: white !important;
}

/* =========================
   卡片式目錄 (Index.html專用) - 動畫收折與大字體
   ========================= */
.card-grid {
  display: grid !important;
  grid-template-columns: repeat(auto-fill, minmax(400px, 1fr)) !important;
  gap: 30px !important;
  padding: 20px 0 50px 0 !important;
}

.card {
  background: var(--article-bg, #ffffff) !important;
  border-radius: var(--radius, 14px) !important;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05) !important;
  padding: 30px !important;
  transition: transform 0.3s ease, box-shadow 0.3s ease !important;
  display: flex !important;
  flex-direction: column !important;
  position: relative !important;
}

.card:hover {
  transform: translateY(-8px) !important;
  box-shadow: 0 16px 40px rgba(0, 0, 0, 0.15) !important;
}

.card-title {
  margin: 0 !important;
  font-family: "Noto Sans TC", serif !important;
  font-size: 24pt !important;
  color: #333 !important;
  line-height: 1.5 !important;
  padding-bottom: 20px !important;
  transition: color 0.3s !important;
  cursor: pointer !important;
}

.card-links {
  display: flex !important;
  flex-wrap: wrap !important;
  gap: 15px !important;
  max-height: 0 !important;
  opacity: 0 !important;
  overflow: hidden !important;
  margin-top: 0 !important;
  transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1) !important;
  pointer-events: none;
}

.card:hover .card-title {
  color: #0066cc !important;
}

.card:hover .card-links {
  max-height: 800px !important;
  opacity: 1 !important;
  margin-top: 20px !important;
  border-top: 2px solid #e0e0e0 !important;
  padding-top: 20px !important;
  pointer-events: auto;
}

.card-links .badge {
  background: #f0f4f8 !important;
  color: #0b57d0 !important;
  padding: 10px 20px !important;
  border-radius: 30px !important;
  font-size: 20pt !important;
  text-decoration: none !important;
  font-weight: 500 !important;
  font-family: "Noto Sans TC", sans-serif !important;
  border: 2px solid transparent !important;
  transition: all 0.2s ease !important;
  display: inline-block;
}

.card-links .badge:hover {
  background: #0b57d0 !important;
  color: #fff !important;
  border-color: #0b57d0 !important;
  box-shadow: 0 4px 12px rgba(11,87,208,0.3) !important;
}
</style>
'''

html = html.replace('</head>', style_block + '</head>')
with open('docs/_template.html', 'w', encoding='utf-8') as f:
    f.write(html)
