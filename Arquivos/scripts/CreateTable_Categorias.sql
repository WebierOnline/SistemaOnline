IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('Categorias') AND type = 'U')
BEGIN
    CREATE TABLE Categorias (
        ID_Categoria INT IDENTITY(1,1) PRIMARY KEY,
        Categoria    VARCHAR(50) NOT NULL,
        Tipo_Empresa INT         NOT NULL
    );
END
GO
