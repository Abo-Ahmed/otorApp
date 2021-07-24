SELECT [ORDER].[ID] , [STMT].[AMNT] 
    FROM [ORDER] 
    LEFT JOIN
            (
                SELECT  [CONTENT].[ORDER_ID] AS ID,
                        IIF(SUM ([CONTENT].[AMOUNT]) IS NULL,0,SUM ([CONTENT].[AMOUNT])) AS AMNT 
                FROM    [CONTENT] 
                WHERE       [CONTENT].[TYPE] = "EJAZA" 
                GROUP BY    [CONTENT].[ORDER_ID]
            ) AS STMT  
    ON [ORDER].[ID] = [STMT].[ID]