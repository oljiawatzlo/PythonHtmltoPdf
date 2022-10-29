import imgkit
from jinja2 import Environment, FileSystemLoader
import pandas as pd
#from weasyprint import HTML, CSS
#import os


con = imgkit.config(wkhtmltoimage = r'wkhtmltopdf\bin\wkhtmltoimage.exe')
#os.add_dll_directory(r'GTK3-Runtime Win64\bin')

df = pd.read_excel('D:/Python Playground/PytonJinja2/game.xlsx', sheet_name='Sheet1', dtype={'Name': str, 'Value': float})
df.rename({"Round" : "Price(Baht)"}, axis=1, inplace=True)
#df.set_index(["Name","Price(Baht)"], inplace=True)
df['ID'] = "images/" + df['ID'].astype(str) + ".jpg"
df['status'] = df['status'].astype(bool)
df = df.reset_index()


# 2. Create a template Environment
env = Environment(loader=FileSystemLoader('templates'))


for i in range(round(len(df)/6)):
#print(df.iloc[0*(i):6*(i),1:])

# 3. Load the template from the Environment
  template = env.get_template('sellgamepic.html')

  try:
    # 4. Render the template with variables
    html = template.render(page_title_text='Facebook Picture',
                            title_text='Key Game Discount',
                            data=df.iloc[0+(i*6):6+(i*6),1:])
  except:
    html = template.render(page_title_text='Facebook Picture',
                            title_text='Key Game Discount',
                            data=df.iloc[0+(i*6):,1:])

  # 5. Write the template to an HTML file
  with open('picture.html', 'w') as f:
        f.write(html)

#css = CSS(string=''' @page {size: 1200px 1200px; margin: 0px; border: 0px} ''')


# HTML('picture.html').write_pdf('report.pdf', stylesheets=[css])

  options = {
      "enable-local-file-access": None
    }

  imgkit.from_file('picture.html', r'output\output'+ "{0}".format(i) +'.jpg', config=con, options=options)
