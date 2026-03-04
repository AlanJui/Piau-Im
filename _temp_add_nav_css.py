css_append = '''
/* Article Nav 樣式 */
.main-nav {
  display: flex !important;
  justify-content: center;
  align-items: center;
  margin-bottom: 40px;
  margin-top: 40px;
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

.main-nav ul li {
  margin: 0;
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
'''
with open('docs/assets/styles/styles.css', 'a', encoding='utf-8') as f:
    f.write(css_append)
print("CSS appended.")
