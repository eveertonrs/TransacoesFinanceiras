--Criação das Tabelas

CREATE TABLE tbdTransacoes (
    Id_Transacao INT IDENTITY(1,1) PRIMARY KEY,   
    Numero_Cartao VARCHAR(16) NOT NULL,           
    Valor_Transacao DECIMAL(10, 2) NOT NULL,      
    Data_Transacao DATETIME NOT NULL,             
    Descricao VARCHAR(255)                        
);

---Procedure

CREATE PROCEDURE sp_CalcularTransacoesPorPeriodo
    @Data_Inicial DATE,  
    @Data_Final DATE     
AS
BEGIN
    SELECT 
        Numero_Cartao,
        SUM(Valor_Transacao) AS Valor_Total,      
        COUNT(*) AS Quantidade_Transacoes         
    FROM 
        tbdTransacoes                            
    WHERE 
        Data_Transacao BETWEEN @Data_Inicial AND @Data_Final 
    GROUP BY 
        Numero_Cartao                            
    ORDER BY 
        Numero_Cartao                             
END


--FUNCTION

CREATE FUNCTION dbo.CategorizarTransacao(@ValorTransacao DECIMAL(10, 2))
RETURNS VARCHAR(10)
AS
BEGIN
    DECLARE @Categoria VARCHAR(10);

    IF @ValorTransacao > 1000
        SET @Categoria = 'Alta';
    ELSE IF @ValorTransacao >= 500 AND @ValorTransacao <= 1000
        SET @Categoria = 'Média';
    ELSE
        SET @Categoria = 'Baixa';

    RETURN @Categoria;
END;

--Usando a FUNCTION na consulta
SELECT 
    Id_Transacao,
    Numero_Cartao,
    Valor_Transacao,
    Data_Transacao,
    Descricao,
    dbo.CategorizarTransacao(Valor_Transacao) AS Categoria
FROM 
    tbdTransacoes;
	
	
--tabela de Clientes
CREATE TABLE Clientes (
    Id_Cliente INT IDENTITY(1,1) PRIMARY KEY,
    Nome_Cliente VARCHAR(100) NOT NULL,
    Numero_Cartao VARCHAR(16) NOT NULL
);

---VIEW ( No contexto do CRUD que foi solicitado, não foi mencionado a criação da tabela Clientes, nem foi indicado que a tabela tbdTransacoes possui informações relacionadas a clientes. Portanto, a VIEW apresentada é meramente ilustrativa e demonstra como criar uma VIEW que combinaria as informações de duas tabelas, caso elas existissem e fossem corretamente relacionadas.)

CREATE VIEW vw_TransacoesComClientes AS
SELECT 
    c.Nome_Cliente,
    t.Numero_Cartao,
    t.Valor_Transacao,
    t.Data_Transacao,
    dbo.CategorizarTransacao(t.Valor_Transacao) AS Categoria
FROM 
    tbdTransacoes t
JOIN 
    Clientes c ON t.Id_Cliente = c.Id_Cliente; 

--consulta da view
SELECT * FROM vw_TransacoesComClientes;
