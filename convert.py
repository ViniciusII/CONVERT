from numpy.core.numeric import NaN
import pandas as pd
import xlsxwriter
import csv
import re
import os

use_sku_img = False
use_produto_pai = False
descricao_title = False

try:
    #Create target Directory
    os.mkdir('new')
    print("Directory 'new' Created ") 
except FileExistsError:
    print("Directory 'new' already exists")
    

with open('imagens.csv', 'r', encoding='utf-8') as f:
    img = list(csv.reader(f))

sku_file = pd.read_csv('sku.csv', converters={i: str for i in range(0, 30)})

categ_cadastrada = []
marcas = []
att_cadastrados = []

workbook = xlsxwriter.Workbook('new/produtos.xlsx')

worksheet_p = workbook.add_worksheet(name = "produtos")
row_p = 0
worksheet_p.write(row_p, 0,  "produto_sku")  #A
worksheet_p.write(row_p, 1,  "produto_nome") #B
worksheet_p.write(row_p, 2,  "sku") #C
worksheet_p.write(row_p, 3,  "ean") #D
worksheet_p.write(row_p, 4,  "nome") #E
worksheet_p.write(row_p, 5,  "marca") #F
worksheet_p.write(row_p, 6,  "categoria") #G
worksheet_p.write(row_p, 7,  "origem") #H
worksheet_p.write(row_p, 8,  "garantia_dias") #I
worksheet_p.write(row_p, 9,  "peso_liquido_kg") #J
worksheet_p.write(row_p, 10, "peso_cubado_kg") #H
worksheet_p.write(row_p, 11, "altura_produto_cm") #L
worksheet_p.write(row_p, 12, "largura_produto_cm") #M
worksheet_p.write(row_p, 13, "comprimento_produto_cm") #N
worksheet_p.write(row_p, 14, "altura_embalagem_cm") #O
worksheet_p.write(row_p, 15, "largura_embalagem_cm") #P
worksheet_p.write(row_p, 16, "comprimento_embalagem_cm") #Q
worksheet_p.write(row_p, 17, "descricao") #R
row_p = 1


worksheet_a = workbook.add_worksheet(name = "atributos")
row_a = 0
worksheet_a.write(row_a, 0,  "sku")  #A
worksheet_a.write(row_a, 1,  "nome") #B
worksheet_a.write(row_a, 2,  "valor") #C
row_a = 1


worksheet_i = workbook.add_worksheet(name = "imagens")
row_i = 0
worksheet_i.write(row_i, 0,  "sku")  #A
worksheet_i.write(row_i, 1,  "ordem") #B
worksheet_i.write(row_i, 2,  "url") #C
row_i = 1

workMarca = xlsxwriter.Workbook('new/marcas_de_para.xlsx')

worksheet_m = workMarca.add_worksheet(name = "marcas")
row_m = 0
worksheet_m.write(row_m, 0,  "napphub_nome")  #A
worksheet_m.write(row_m, 1,  "externo_id") #B
worksheet_m.write(row_m, 2,  "externo_nome") #C
row_m = 1

workCat = xlsxwriter.Workbook('new/categorias_de_para.xlsx')

worksheet_c = workCat.add_worksheet(name = "categorias")
row_c = 0
worksheet_c.write(row_c, 0,  "napphub_nome")  #A
worksheet_c.write(row_c, 1,  "externo_id") #B
worksheet_c.write(row_c, 2,  "externo_nome") #C
row_c = 1


workAtt = xlsxwriter.Workbook('new/atributos_de_para.xlsx')

worksheet_d = workAtt.add_worksheet(name = "atributos")
row_d = 0
worksheet_d.write(row_d, 0,  "napphub_nome")  #A
worksheet_d.write(row_d, 1,  "napphub_valor") #B
worksheet_d.write(row_d, 2,  "externo_nome_id") #C
worksheet_d.write(row_d, 3,  "externo_nome") #D
worksheet_d.write(row_d, 4,  "externo_valor_id") #E
worksheet_d.write(row_d, 5,  "externo_valor") #F
worksheet_d.write(row_d, 6,  "externo_categoria_id") #G
row_d = 1

produtos_duplicados = sku_file.duplicated(subset=['product_id'], keep=False)

for index, p in sku_file.iterrows():
    try:
        ean = str(int(p['ean']))
    except:
        ean = ''
    product_id = str(p['product_id']).strip()
    sku = str(p['sku']).strip()
    try:
        tam = str(p['tamanho']).strip()
    except KeyError:
        tam = ''
    nome = str(p['nome_sku']).title()
    produto = str(p['nome_produto']).title()
    if use_produto_pai:
        is_prod = bool(produtos_duplicados.iloc[index])
        if is_prod:
            worksheet_p.write(row_p, 0, product_id)  #A
            worksheet_p.write(row_p, 1, produto) #B
        else:
            worksheet_p.write(row_p, 0, '')  #A
            worksheet_p.write(row_p, 1, '') #B
    else:
        worksheet_p.write(row_p, 0, product_id)  #A
        worksheet_p.write(row_p, 1, produto) #B
    
    worksheet_p.write(row_p, 2, sku) #C
    worksheet_p.write(row_p, 3, ean) #D
    if tam == '':
        worksheet_p.write(row_p, 4, nome) #E
    else:
        nome = nome.replace(f' {tam.title()}', f' {tam}')
        worksheet_p.write(row_p, 4, nome) #E
    worksheet_p.write(row_p, 4, nome) #E
    worksheet_p.write(row_p, 5, p['marca'].title()) #F
    worksheet_p.write(row_p, 6, str(p['categoria_hub'])) #G
    worksheet_p.write(row_p, 7, "") #H
    worksheet_p.write(row_p, 8, "") #I
    worksheet_p.write(row_p, 9,  "") #J
    worksheet_p.write(row_p, 10, "") #H
    worksheet_p.write(row_p, 11, "") #L
    worksheet_p.write(row_p, 12, "") #M
    worksheet_p.write(row_p, 13, "") #N
    worksheet_p.write(row_p, 14, "") #O
    worksheet_p.write(row_p, 15, "") #P
    worksheet_p.write(row_p, 16, "") #Q
    descricao = str(p['descricao'])
    if descricao != 'nan' and descricao != '':
        if descricao_title:
            if tam == '':
                worksheet_p.write(row_p, 17, descricao.title()) #R
            else:
                descricao = descricao.title()
                descricao = descricao.replace(f' {tam.title()}', f' {tam}')
                worksheet_p.write(row_p, 17, descricao) #R
        else:
            worksheet_p.write(row_p, 17, descricao) #R
    else:
        worksheet_p.write(row_p, 17, produto) #R
    row_p += 1
    
    cat = str(int(p['categoria']))
    try:
        try:
            tamanho_id = str(int(p['tamanho_id']))
        except ValueError:
            tamanho_id = ''

        if tam != "nan" and tamanho_id != "":
            worksheet_a.write(row_a, 0, sku)#A
            worksheet_a.write(row_a, 1,  "Tamanho") #B
            worksheet_a.write(row_a, 2,  tam) #C
            row_a +=1
            
            tamanho_atual = [tam, tamanho_id, cat]
            if tamanho_atual in att_cadastrados:
                pass
            else:
                worksheet_d.write(row_d, 0, "Tamanho")#A
                worksheet_d.write(row_d, 1, tam)#B
                worksheet_d.write(row_d, 2, tamanho_id)#C
                worksheet_d.write(row_d, 3, "Tamanho")#D
                worksheet_d.write(row_d, 4, "")#E
                worksheet_d.write(row_d, 5, tam)#F
                worksheet_d.write(row_d, 6, cat)#G
                row_d +=1
                att_cadastrados.append(tamanho_atual)
        else:
            pass
    except KeyError:
        pass
    except ValueError:
        pass
    
    try:
        cor = str(p['cor']).strip()
        try:
            cor_id = str(int(p['cor_id']))
        except ValueError:
            cor_id = ''
        if cor != "nan" and cor_id !="":
            worksheet_a.write(row_a, 0, sku)#A
            worksheet_a.write(row_a, 1,  "Cor") #B
            worksheet_a.write(row_a, 2,  cor.split(',')[0]) #C
            row_a +=1
            
            cor_atual = [cor, cor_id, cat]
            if cor_atual in att_cadastrados:
                pass
            else:
                worksheet_d.write(row_d, 0, "Cor")#A
                worksheet_d.write(row_d, 1, cor.split(',')[0])#B
                worksheet_d.write(row_d, 2, cor_id)#C
                worksheet_d.write(row_d, 3, "Cor")#D
                worksheet_d.write(row_d, 4, "")#E
                worksheet_d.write(row_d, 5, cor)#F
                worksheet_d.write(row_d, 6, cat)#G
                row_d +=1
                att_cadastrados.append(cor_atual)
        else:
            pass
    except KeyError:
        pass
    except ValueError:
        pass

    try:
        volume = str(p['volume']).strip()
        try:
            volume_id = str(int(p['volume_id']))
        except ValueError:
            volume_id = ''
        if volume != "nan" and volume_id !="":
            worksheet_a.write(row_a, 0, sku)#A
            worksheet_a.write(row_a, 1,  "Volume") #B
            worksheet_a.write(row_a, 2,  volume.split(',')[0]) #C
            row_a +=1
            
            volume_atual = [volume, volume_id, cat]
            if volume_atual in att_cadastrados:
                pass
            else:
                worksheet_d.write(row_d, 0, "Volume")#A
                worksheet_d.write(row_d, 1, volume.split(',')[0])#B
                worksheet_d.write(row_d, 2, volume_id)#C
                worksheet_d.write(row_d, 3, "Volume")#D
                worksheet_d.write(row_d, 4, "")#E
                worksheet_d.write(row_d, 5, volume)#F
                worksheet_d.write(row_d, 6, cat)#G
                row_d +=1
                att_cadastrados.append(volume_atual)
        else:
            pass
    except KeyError:
        pass
    except ValueError:
        pass

    try:
        voltagem = str(p['voltagem']).strip()
        try:
            voltagem_id = str(int(p['voltagem_id']))
        except ValueError:
            voltagem_id = ''
        if voltagem != "nan" and voltagem_id !="":
            worksheet_a.write(row_a, 0, sku)#A
            worksheet_a.write(row_a, 1,  "Voltagem") #B
            worksheet_a.write(row_a, 2,  voltagem.split(',')[0]) #C
            row_a +=1
            
            voltagem_atual = [voltagem, voltagem_id, cat]
            if voltagem_atual in att_cadastrados:
                pass
            else:
                worksheet_d.write(row_d, 0, "Voltagem")#A
                worksheet_d.write(row_d, 1, voltagem.split(',')[0])#B
                worksheet_d.write(row_d, 2, voltagem_id)#C
                worksheet_d.write(row_d, 3, "Voltagem")#D
                worksheet_d.write(row_d, 4, "")#E
                worksheet_d.write(row_d, 5, voltagem)#F
                worksheet_d.write(row_d, 6, cat)#G
                row_d +=1
                att_cadastrados.append(voltagem_atual)
        else:
            pass
    except KeyError:
        pass
    except ValueError:
        pass

    try:
        capacidade = str(p['capacidade']).strip()
        try:
            capacidade_id = str(int(p['capacidade_id']))
        except ValueError:
            capacidade_id = ''
        if capacidade != "nan" and capacidade_id !="":
            worksheet_a.write(row_a, 0, sku)#A
            worksheet_a.write(row_a, 1,  "Capacidade") #B
            worksheet_a.write(row_a, 2,  capacidade.split(',')[0]) #C
            row_a +=1
            
            capacidade_atual = [capacidade, capacidade_id, cat]
            if capacidade_atual in att_cadastrados:
                pass
            else:
                worksheet_d.write(row_d, 0, "Capacidade")#A
                worksheet_d.write(row_d, 1, capacidade.split(',')[0])#B
                worksheet_d.write(row_d, 2, capacidade_id)#C
                worksheet_d.write(row_d, 3, "Capacidade")#D
                worksheet_d.write(row_d, 4, "")#E
                worksheet_d.write(row_d, 5, capacidade)#F
                worksheet_d.write(row_d, 6, cat)#G
                row_d +=1
                att_cadastrados.append(capacidade_atual)
        else:
            pass
    except KeyError:
        pass
    except ValueError:
        pass

    try:
        composicao = str(p['composicao']).strip()
        try:
            composicao_id = str(int(p['composicao_id']))
        except ValueError:
            composicao_id = ''
        if composicao != "nan" and composicao_id !="":
            worksheet_a.write(row_a, 0, sku)#A
            worksheet_a.write(row_a, 1,  "Composicao") #B
            worksheet_a.write(row_a, 2,  composicao.split(',')[0]) #C
            row_a +=1
            
            composicao_atual = [composicao, composicao_id, cat]
            if composicao_atual in att_cadastrados:
                pass
            else:
                worksheet_d.write(row_d, 0, "Composicao")#A
                worksheet_d.write(row_d, 1, composicao.split(',')[0])#B
                worksheet_d.write(row_d, 2, composicao_id)#C
                worksheet_d.write(row_d, 3, "Composicao")#D
                worksheet_d.write(row_d, 4, "")#E
                worksheet_d.write(row_d, 5, composicao)#F
                worksheet_d.write(row_d, 6, cat)#G
                row_d +=1
                att_cadastrados.append(composicao_atual)
        else:
            pass
    except KeyError:
        pass
    except ValueError:
        pass

    for i in range(len(img)):
        img_code = str(img[i][0]).strip()
        if use_sku_img:
            ref_code = sku
        else:
            ref_code = product_id

        if ref_code == img_code:
            for j in range(len(img[i])):
                iter_item = str(img[i][j]).strip()
                if "http" in iter_item:
                    worksheet_i.write(row_i, 0, sku) #A
                    worksheet_i.write(row_i, 1, j) #B
                    worksheet_i.write(row_i, 2, iter_item) #C
                    row_i +=1
                    
    categ_hub = p['categoria_hub']
    categ_nome = p['categoria_nome']
    categ_atual = [categ_hub,cat,categ_nome]
    if categ_atual in categ_cadastrada:
        pass
    else:
        worksheet_c.write(row_c, 0, categ_hub) #A
        worksheet_c.write(row_c, 1, cat) #B
        worksheet_c.write(row_c, 2, categ_nome) #C
        row_c += 1
        categ_cadastrada.append(categ_atual)
        
    marca = p['marca'].title()
    marca_id = str(int(p['marca_id']))
    marca_atual = [marca, marca_id]
    if marca_atual in marcas:
        pass
    else:
        worksheet_m.write(row_m, 0, marca)#A
        worksheet_m.write(row_m, 1, marca_id)#A
        worksheet_m.write(row_m, 2, marca)#A
        row_m += 1
        marcas.append(marca_atual)
        
workbook.close()
workMarca.close()
workCat.close()
workAtt.close()

