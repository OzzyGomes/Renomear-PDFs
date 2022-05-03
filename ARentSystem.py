#!/usr/bin/env python
# coding: utf-8

# In[1]:


from PyPDF4 import PdfFileReader, PdfFileWriter
import pandas as pd
import re
import os
import glob
import tabula
import pytesseract as tess
from PIL import Image
import fitz
import io
import pdfquery
from pdf2image import convert_from_path
import textract

tess.pytesseract.tesseract_cmd = r'C:\Users\ojgomes\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'


# # O código começa aqui

# In[2]:



def pdf_to_text(doc):
    images = convert_from_path(doc)


    texto = tess.image_to_string(images[0])
    #for i in range(len(images)):
   
    # Save pages as images in the pdf
    #images[i].save('page'+ str(i) +'.jpg', 'JPEG')

    #texto = tess.image_to_string(images[i])
    #break
    
    return texto


# In[3]:



#função para extrair imagem do pdf e em seguida retirar texto da imagem 
def extract_text_from_img_doc(doc):
    
    file = fitz.open(doc)
    
    for page_index in range(len(file)):
        #get the page itself
        page = file[page_index]
        image_list = page.getImageList()

        #print number of image this page
        if image_list:
            print(f'[+] Found a total of {len(image_list)} imagesin page {page_index}')
        else:
            print('[!] No image found on page', page_index)
        for image_index, img in enumerate(page.getImageList(), start = 1):
            #get the XREF of the image
            xref = img[0]

            #extract image bytes
            base_image = file.extractImage(xref)
            image_bytes = base_image['image']

            #get the image extension
            image_ext = base_image['ext']

            #load it to PIL
            image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
            
            #save it in local disk
            
            #image.save(open(f'image{page_index+1}_{image_index}.{image_ext}', 'wb'))
    file.close()
    texto = tess.image_to_string(image)
    return texto


# # Procurando apenas por CNPJ

# In[4]:


#ler pasta onde ficam os boletos
file_path = 'boletos/'

#transformando df para uma lista de todos os pdfs da pasta com a bliblioteca glob
df = glob.glob(os.path.join(file_path,'*.pdf'))

df2 = pd.read_excel('Siglas_CNPJs.xlsx')

text = ''

master_list = []

# com df como lista de pdfs itero cada pdf com for abrindo e adicionando as funções que desejo
for i in df:
    
    KeyW3_4 = ''
    
    #leio o pdf com pdf4, informo a quantidade de paginas e em seguida digo que quero o texto do conteudo
    pdf = PdfFileReader(i)
    obtrpage= pdf.getPage(0)
    txt = obtrpage.extractText()
    
    #encontrar CNPJ com regular expression
    KeyW4 = re.findall('\d{2}.\d{3}.\d{3}/\d{4}-\d{2}', txt)
    
    
    #caso o pdf não tenha texto e sim uma imagem não lerá PDF4, assim crie uma condição para ler imagem da msm iteração
    if len(KeyW4) == 0:
        try:
            ban_texto = pdf_to_text(i)
            KeyW4 = re.findall('\d{2}.\d{3}.\d{3}/\d{4}-\d{2}', ban_texto)
            
        except:
            print(f'erro ao ler com função pdf_to_text, boleto:  {i}')
    
    
    
    # em outra condição é tbm para imagems, mas com outra biblioteca para caso a primeira falhar
    if len(KeyW4) == 0:
        try:
            texto = extract_text_from_img_doc(i)
            KeyW4 = re.findall('\d{2}.\d{3}.\d{3}/\d{4}-\d{2}', texto)
                
        except:
            print('nenhuma imagem encontrada')
        
    #caso ainda o PDF não leia outra excelente biblioteca é tabula que ler em formato de tabelas 
    if len(KeyW4) == 0:
        try:
            #ler o pdf, converte de dataframe para lista, da lista converte para string e finalmente obtem o CNPJ com re
            pdft = tabula.read_pdf(i)
            
            pdf_list = pdft[0].iloc[:, 0].values.tolist()
            
            srt_tex = "".join(map(str, pdf_list))
            
            KeyW4 = re.findall('\d{2}.\d{3}.\d{3}/\d{4}-\d{2}', srt_tex)
            
        except:
            pass
    else:
        pass
            
    #removendo duplicidade das listas
    KeyW4 = list(dict.fromkeys(KeyW4))
    
    #removendo o item abaixo
    try:
        KeyW4.remove('42.591.651/0001-43')
    except:
        pass
    
    #KeyW4 = ','.join(KeyW4)
    
    key = [KeyW4]
    
    #este loop varre a lista key e se o item for diferente de 0 então ele aplica o Try 
    for j in key:
        if len(j)!=0:
            try:
                #tenta através de Pandas encontrar o arquivo
                search = df2[df2['CNPJ'].str.contains(j[0], case=False, na=False) | df2['CNPJ 2'].str.contains(j[0], case=False, na=False)]
                sigla = search['SIGLA'].iloc[0] + ' - ' + search['NOVO CC'].iloc[0]
                
                if len(sigla)<=12:
                    print(sigla)
                    os.rename(i, 'boletos\\' + sigla + ' - ' + os.path.basename(i))
                    break

                else:
                    print('o conteudo de sigla não é igual a 12')
            except:
                print(f'erro ao procurar keyword "{j}"')
        else:
            print('erro ')
    

    
    
    
    


# In[ ]:




