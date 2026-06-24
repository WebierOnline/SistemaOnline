"""
patch_cmunfgibs.py

GeraIdentificacao ja tem o parametro cMunFGIBS As Long (confirmado na imagem57).
Atualmente passa 0. Para operacoes interestaduais (IdentificadorDestino=2),
deve passar o CodigoIBGE do destinatario — ja disponivel em Destinatario!CodigoIBGE
porque o recordset e aberto com SELECT * FROM cliente/fornecedor.

P1: insere calculo de lCMunFGIBS antes da chamada GeraIdentificacao
P2: substitui o 0 hardcoded por lCMunFGIBS na chamada
"""
import sys, shutil

BAS = r"C:\Projeto\Compartilhado\Modulos\modNFe.bas"
BAK = BAS + ".bak_cmunfgibs"

shutil.copy2(BAS, BAK)
print(f"Backup: {BAK}")

with open(BAS, "rb") as f:
    data = f.read()

def norm(b):
    return b.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")

# P1: inserir calculo de lCMunFGIBS antes de dhEmi
OLD1 = b"    dhEmi = Format(NFe!DataEmissao, \"yyyy/mm/dd\")\r\n"
NEW1 = (
    b"    dhEmi = Format(NFe!DataEmissao, \"yyyy/mm/dd\")\r\n"
    b"\r\n"
    b"    Dim lCMunFGIBS As Long\r\n"
    b"    lCMunFGIBS = 0\r\n"
    b"    If Val(NFe!IdentificadorDestino) = 2 Then\r\n"
    b"        lCMunFGIBS = IIf(IsNull(Destinatario!CodigoIBGE), 0, CLng(Val(Destinatario!CodigoIBGE)))\r\n"
    b"    End If\r\n"
)

# P2: substituir cMunFGIBS=0 apenas na chamada NFe (sistNFe, nao sistNFCe)
# Usa prefixo unico "sistNFe.GeraIdentificacao" para diferenciar da linha NFCe
OLD2 = b"    iRetorno = sistNFe.GeraIdentificacao(NFe!cCodigoNota, NFe!NaturezaOperacao, 55, NFe!SerieNF, NFe!NumeroNota, Format(NFe!DataEmissao, \"yyyy-mm-dd\") & \"T\" & Format(Time, \"hh:mm:ss\") & UTC, Format(NFe!DataSaida, \"yyyy-mm-dd\") & \"T\" & Format(Time, \"hh:mm:ss\") & UTC, NFe!TipoDocumento, 0, NFe!IdentificadorDestino, Parametros!CodigoIBGE, 1, Left(NFe!FormatoEmissaoNFe, 1), Left(NFe!FinalidadeEmissaoNFe, 1), indFinal, 1, Left$(\"ONLINE COMMERCE - v.\" & CStr(App.Major) & \".\" & CStr(App.Minor) & \".\" & CStr(App.Revision), 20), dhContingencia, justContingencia, 0, 0, \"\", \"\", mensagemAlerta, mensagemErro)\r\n"
NEW2 = b"    iRetorno = sistNFe.GeraIdentificacao(NFe!cCodigoNota, NFe!NaturezaOperacao, 55, NFe!SerieNF, NFe!NumeroNota, Format(NFe!DataEmissao, \"yyyy-mm-dd\") & \"T\" & Format(Time, \"hh:mm:ss\") & UTC, Format(NFe!DataSaida, \"yyyy-mm-dd\") & \"T\" & Format(Time, \"hh:mm:ss\") & UTC, NFe!TipoDocumento, 0, NFe!IdentificadorDestino, Parametros!CodigoIBGE, 1, Left(NFe!FormatoEmissaoNFe, 1), Left(NFe!FinalidadeEmissaoNFe, 1), indFinal, 1, Left$(\"ONLINE COMMERCE - v.\" & CStr(App.Major) & \".\" & CStr(App.Minor) & \".\" & CStr(App.Revision), 20), dhContingencia, justContingencia, 0, lCMunFGIBS, \"\", \"\", mensagemAlerta, mensagemErro)\r\n"

patches = [(OLD1, NEW1, "P1: lCMunFGIBS antes de dhEmi"),
           (OLD2, NEW2, "P2: cMunFGIBS na chamada GeraIdentificacao")]

errors = 0
for old, new, desc in patches:
    cnt = data.count(old)
    if cnt != 1:
        print(f"ERRO {desc}: count={cnt} (esperado 1)")
        errors += 1
    else:
        data = data.replace(old, new)
        print(f"OK   {desc}")

data = norm(data)

if errors:
    print(f"\n{errors} erro(s). Arquivo NAO foi salvo.")
    sys.exit(1)

with open(BAS, "wb") as f:
    f.write(data)
print("\nArquivo salvo com sucesso.")
