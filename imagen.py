#Desarrollado por Alejandra López Vega
#Para Comercializadora de Plásticos de Alta Calidad S.A. de C.V.
#Version 1.0 lanzada en Octubre de 2025
from PIL import Image, ImageDraw
img = Image.open('blank.png')
draw_img = ImageDraw.Draw(img)
 
data = ['4','5','87','1','44','83','93','2','54','84','100','64'] 
x = 0
 
for i in data:
    x = x + 30 
    y = 200 - int(i) 
    draw_img.line((x,200,x,y), width=10, fill=(255,0,0,255))
         
img.show()