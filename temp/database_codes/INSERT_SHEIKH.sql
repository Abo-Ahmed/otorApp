
--------------------------------------------------------
--  DDL for Procedure INSERT_SHEIKH
--------------------------------------------------------

  CREATE OR REPLACE PROCEDURE "OTOR"."INSERT_SHEIKH" (                
                                    P_FUN_ID                   IN        VARCHAR2,
                                    P_MSG_ID                   IN        VARCHAR2,                                                 
                                    P_STATUS_CODE              OUT       VARCHAR2,
                                    P_STATUS_DESC              OUT       VARCHAR2,
                                    -----------------------------------
                                    P_NAME                     IN    VARCHAR2,
                                    P_PHONE_1                  IN    VARCHAR2,--
                                    P_PHONE_2                  IN    VARCHAR2,--
                                    P_CITY                     IN    VARCHAR2,
                                    P_COUNTRY                  IN    VARCHAR2,
                                    P_ADDRESS                  IN    VARCHAR2
                                )
IS 

 
 
 BEGIN
 
    P_STATUS_DESC := "INTIAL STATE";
    
    IF P_NAME IS NULL THEN 
       P_STATUS_DESC :=  "NAME IS REQUIRED";
    END IF; 
    
    IF P_PHONE_1 IS NULL THEN 
       P_STATUS_DESC :=  "PHONE IS REQUIRED";
    END IF; 
            
    INSERT INTO SHEIKH(   ID,
                          NAME ,
                          PHONE_1,
                          PHONE_2,
                          COUNTRY,
                          CITY,
                          ADDRESS
                          )
                  VALUES( ID,
                          P_NAME ,
                          P_PHONE_1,
                          P_PHONE_2,
                          P_COUNTRY,
                          P_CITY,
                          P_ADDRESS
                        );
      END IF;
   
END INSERT_SHEIKH;

