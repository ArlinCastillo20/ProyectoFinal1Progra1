CREATE TABLE Guisel_Asistente(
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Prompt NVARCHAR(MAX),
    Resultado NVARCHAR(MAX),
    FechaHora DATETIME DEFAULT GETDATE()
);