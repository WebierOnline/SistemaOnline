# -*- coding: utf-8 -*-
"""Cria AlterTable_Categorias_Padrao.sql (ASCII/CRLF), adiciona linha 111 no _manifesto.txt
   e espelha os dois em C:\\scripts (byte-a-byte)."""
import shutil

SQL_NAME = "AlterTable_Categorias_Padrao.sql"
REPO = r"C:\projeto\Arquivos\scripts"
MIRROR = r"C:\scripts"

sql = (
    "-- Categorias.Padrao: marca a categoria padrao daquele Tipo_Empresa.\r\n"
    "-- Usado por Produtos_Cadastro (produto novo ja vem com a categoria padrao selecionada)\r\n"
    "-- e definido em Categorias_Cadastro pelo botao \"Padrao\".\r\n"
    "-- A regra \"so uma padrao por Tipo_Empresa\" e garantida pela aplicacao (o UPDATE zera as\r\n"
    "-- demais antes de marcar a nova), NAO por indice filtrado - indice filtrado quebra\r\n"
    "-- INSERT/UPDATE via Provider=SQLOLEDB no VB6 (ARITHABORT OFF).\r\n"
    "-- Idempotente.\r\n"
    "SET NOCOUNT ON;\r\n"
    "GO\r\n"
    "\r\n"
    "IF NOT EXISTS (\r\n"
    "    SELECT 1 FROM sys.columns\r\n"
    "    WHERE object_id = OBJECT_ID('Categorias') AND name = 'Padrao'\r\n"
    ")\r\n"
    "    ALTER TABLE Categorias\r\n"
    "        ADD Padrao BIT NOT NULL CONSTRAINT DF_Categorias_Padrao DEFAULT (0);\r\n"
    "GO\r\n"
)

for base in (REPO, MIRROR):
    p = base + "\\" + SQL_NAME
    with open(p, "wb") as f:
        f.write(sql.encode("ascii"))
    print("escrito", p)

# manifesto: append linha 111 preservando cp1252 + CRLF
for base in (REPO, MIRROR):
    p = base + "\\_manifesto.txt"
    with open(p, "rb") as f:
        data = f.read()
    if b"111|" + SQL_NAME.encode() in data:
        print("manifesto ja tem 111:", p)
        continue
    if not data.endswith(b"\r\n"):
        data += b"\r\n"
    data += b"111|" + SQL_NAME.encode("ascii") + b"|GERAL\r\n"
    with open(p, "wb") as f:
        f.write(data)
    print("manifesto atualizado", p)

print("Feito.")
