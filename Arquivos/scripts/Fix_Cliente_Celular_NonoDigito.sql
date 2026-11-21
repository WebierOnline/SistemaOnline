-- Celulares antigos de 8 digitos (antes do 9 extra) que ficaram no cadastro como (89) 94275-280_ ou (89) 9427-5280:
-- acrescenta o 9 depois do DDD e grava no formato da mascara: (89) 99427-5280.
-- So mexe em cliente.celular com exatamente 10 digitos cujo primeiro digito depois do DDD e 6, 7, 8 ou 9 (celular).
-- Os de 10 digitos comecando com 2 a 5 sao telefone fixo e ficam como estao. Idempotente (depois da correcao tem 11 digitos).
;WITH x AS (
    SELECT celular,
           REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(celular, '(', ''), ')', ''), '-', ''), ' ', ''), '_', ''), '.', '') AS d
    FROM dbo.cliente
    WHERE celular IS NOT NULL
)
UPDATE x
SET celular = '(' + LEFT(d, 2) + ') 9' + SUBSTRING(d, 3, 4) + '-' + SUBSTRING(d, 7, 4)
WHERE LEN(d) = 10
  AND d NOT LIKE '%[^0-9]%'
  AND SUBSTRING(d, 3, 1) IN ('6', '7', '8', '9')
GO
