-- Tabela de Tags vinculadas a Categorias
IF NOT EXISTS (SELECT 1 FROM sys.objects WHERE object_id = OBJECT_ID('Categorias_Tags') AND type = 'U')
BEGIN
    CREATE TABLE Categorias_Tags (
        ID_Tags     INT IDENTITY(1,1) PRIMARY KEY,
        Tags        NVARCHAR(100) NOT NULL,
        ID_Categoria INT NOT NULL,
        CONSTRAINT FK_CategoriasTags_Categoria
            FOREIGN KEY (ID_Categoria) REFERENCES Categorias(ID_Categoria)
    );
END
