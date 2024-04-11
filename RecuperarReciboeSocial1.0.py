import fdb
import os
from xml.dom import minidom
from pathlib import Path

caminhoPadrao = 'C:/Digimax/MaxFolha/'
caminhoDosArquivos = caminhoPadrao +"pEsocial/"
origem = caminhoPadrao +"ESocial/"
listaDeProtocolos = []
lista = []
listaId = []
municipios = ['054','130','154','141']
tpOrgao = ['P','C','I']
rec = "-rec.xml"
grupo = "-env-lot.xml"

def conectarBanco():
    arquivo = open(caminhoPadrao + "MaxFolha.ini", "r")
    conteudo = arquivo.readlines()
    firebid = conteudo[1].replace("\\","/").replace('caminho=','').strip()
    porta = conteudo[2].replace('porta=','') .strip()
    servidor = conteudo[3].replace('servidor=','').strip()
    base = conteudo[4].replace('base=','').strip()
    
    caminhoCompleto = (servidor +'/'+ porta + ':' + base).replace("\\","/")

    fdb.load_api(firebid)
    conexao = fdb.connect(dsn = caminhoCompleto, user = '', password = '')
    cursor = conexao.cursor()
    SELECT = "select id_geracao_esocial from lote_esocial where id_geracao_esocial = " + loteDefined
    INSERT = "INSERT INTO LOTE_ESOCIAL (ID, ID_GERACAO_ESOCIAL, GRUPO_LOTE, VERSAO_LOTE, XML_LOTE, RETWS_LOTE, ARQ_ENVIO_LOTE, ARQ_RETORNO_LOTE) VALUES (NEXT VALUE FOR lote_esocial_gen, +"+ loteDefined + "," + grupoNumber + ",'10.20', '', '', 'env-lot', 'rec')"

    try:
        for id_geracao_esocial in cursor.execute(SELECT):
            listaId.append(id_geracao_esocial)
            x = len(listaId)
        for x in range(x,qtdRecArquivo):
            cursor.execute(INSERT)
            listaId.append(id_geracao_esocial)
        conexao.commit()
        gerarArquivosTxt()
        print("Linhas geradas: " + str(len(listaId)))
        try:
            os.startfile(caminhoPadrao+ "MaxFolhaEsocial.exe")
        except:
            return
    except:
        try:
            y=0
            for y in range(y,qtdRecArquivo):
                cursor.execute(INSERT)
                conexao.commit()
            print("Linhas geradas: " + str(qtdRecArquivo))
            os.startfile(caminhoPadrao+ "MaxFolhaEsocial.exe")

        except:
            print("Lote " + loteDefined +" não gerado")
    finally:
        print("Encerrando sistema.")
        os.system("PAUSE")

def informarMunLote(erro = False):
    global mun
    global loteName
    global loteDefined
    
    if erro == True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print('Erro! Caminho invalido ou arquivo sem lote.')
        lista.clear();listaDeProtocolos.clear()
        os.system("PAUSE")
    try:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("Selecione o seu municipio:\n 0 - Eusebio \n 1 - Parambu \n 2 - Santana Do Cariri \n 3 - Potengi")
        mun = int(input())
        print("Selecione: \n 0 - Prefeitura \n 1 - Camara \n 2 - Instituto de Previdencia")
        orgao = input()
        print("Qual lote deseja recuperar recibos?")
        lote = input()
        loteDefined = lote
        loteName = tpOrgao[int(orgao)]+"_"+municipios[mun]+"_"+lote
        validarCaminho(loteName)
    except:
        informarMunLote(erro = True)

def validarCaminho(lote):
    global grupoNumber
    global qtdRecArquivo
    pathLote = Path(origem+lote+'/LOTE')
    if pathLote.exists():   
        for f in os.listdir(pathLote):
            if rec in f:
                lista.append(f)
            if grupo in f:
                grupoName = f
        qtdRecArquivo = len(lista)

        os.system('cls' if os.name == 'nt' else 'clear')
        print("Protocolos de recibos do lote: " + lote)
        for i in range(len(lista)):
            with open(origem+lote+"/LOTE/" + lista[i],"r", encoding='utf-8') as f:
                xml = minidom.parse(f)
                protocolo = xml.getElementsByTagName("protocoloEnvio")
            for tag in protocolo:
                print(tag.firstChild.data)
                varAux = tag.firstChild.data
                listaDeProtocolos.append(varAux)

        with open(origem+lote+"/LOTE/" + grupoName,"r", encoding='utf-8') as f:
            xml = minidom.parse(f)
            grupoLote = xml.getElementsByTagName("envioLoteEventos")
            for tag in grupoLote:
                grupoNumber = tag.getAttribute("grupo")
                continue

        print("Deseja incluir a referencia do lote no banco de dados? \n S - SIM  N - NÂO")
        includeDB = input()
        if(includeDB == 's' or includeDB == 'S'):
            conectarBanco()
        else:
            print("Encerrando sistema.")
            os.system("PAUSE")
    else:
        informarMunLote(True)

def esvreverNosArquivosTxt():
    maxFolhaCertificado = Path(caminhoDosArquivos + "maxfolhacertificado.txt")
    maxFolhaPath = Path(caminhoDosArquivos + "maxfolhapath.txt")
    maxFolhaProtocolo = Path(caminhoDosArquivos + "maxfolhaprotocolo.txt")

    maxFolhaCertificado.write_text("\n\n")
    maxFolhaPath.write_text(origem+loteName+"/LOTE")
    maxFolhaProtocolo.write_text("\n".join(listaDeProtocolos))

    print("Arquivos gerados com sucesso!")

def gerarArquivosTxt():
    
    os.system('cls' if os.name == 'nt' else 'clear')
    print("Gerando arquivos...")
    try:
        os.mkdir(caminhoDosArquivos)
        try:
            esvreverNosArquivosTxt()
        except:
            print("Erro na geração dos arquivos")
    except:
        esvreverNosArquivosTxt()
    finally:
        print("Concluindo...")

informarMunLote()