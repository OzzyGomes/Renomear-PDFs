import pytesseract as tess
from pdf2image import convert_from_path
import fitz
from PIL import Image
import io

class Text_from_image:
    def __init__(self):
        pass

    def pdf_to_text(self, doc):
        images = convert_from_path(doc)


        texto = tess.image_to_string(images[0])
        #for i in range(len(images)):
    
        # Save pages as images in the pdf
        #images[i].save('page'+ str(i) +'.jpg', 'JPEG')

        #texto = tess.image_to_string(images[i])
        #break
        
        return texto
    
    def extract_text_from_img_doc(self, doc):
    
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
