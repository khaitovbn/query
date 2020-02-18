USE InfoDB_p
GO

/*
ALTER VIEW KPI.V_KPI_DC_LEVEL
AS
SELECT sc.ID as STORECODE_ID
     , ta.ID as TEMPAREA_ID
     , kd.ID as DC_ID
     , m.ID as MACRO_ID
     , o.CODE as OPERATION_ID
     , c.ID as CATEGORY_ID
     , ut.ID as USE_TYPE_ID
     , null as WORK_SHIFT_ID
     , d.ID  as DISTRICT_ID
FROM [DICT].[USE_TYPE] ut 
LEFT JOIN [KPI].[KPI_DC] kd 
	ON  iif (kd.ID is null,0,1 ) = ut.DC
LEFT JOIN DICT.STORECODE sc 
	ON  iif (sc.ID is null,0,1 ) = ut.STORECODE
LEFT JOIN DICT.TEMPAREA ta 
	ON  iif (ta.ID is null,0,1 ) = ut.TEMPAREA
LEFT JOIN DICT.MACRO m 
	ON  iif (m.ID is null,0,1 ) = ut.MACRO
LEFT JOIN DBO.tbProizv_Operations o 
	ON  iif (o.CODE is null,0,1 ) = ut.OPERATION  
LEFT JOIN DICT.CATEGORY c 
	ON  iif (c.ID is null,0,1 ) = ut.CATEGORY
LEFT JOIN DICT.DISTRICT d 
	ON  iif (d.ID is null,0,1 ) = ut.DISTRICT_FLG



GO


SELECT *
FROM [DICT].[KPI_PERIOD_TYPE]

SELECT *
FROM KPI.KPI_NAME


/*
CREATE TABLE [DICT].[KPI_PERIOD_TYPE](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[KPI_NAME_ID] [int] NOT NULL,
	[PERIOD_TYPE_ID] [int] NOT NULL,
 CONSTRAINT [PK_KPI_PERIOD_TYPE] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)
)

TRUNCATE TABLE  [DICT].[KPI_PERIOD_TYPE]

INSERT [DICT].[KPI_PERIOD_TYPE]
SELECT * from (
                 VALUES (  2 ,760)
                      , (  8 ,760)
                      , ( 21 ,760)
                      , ( 23 ,760)
                      , ( 24 ,760)
                      , ( 26 ,760)
                 
              ) as x([KPI_NAME_ID],[PERIOD_TYPE_ID])

SELECT *
FROM [DICT].[KPI_PERIOD_TYPE]
*/


--select STORECODE_ID, TEMPAREA_ID, DC_ID, MACRO_ID, OPERATION_ID, CATEGORY_ID, USE_TYPE_ID, WORK_SHIFT_ID, DISTRICT_ID from KPI.v_KPI_DC_LEVEL
--except
--select STORECODE_ID, TEMPAREA_ID, DC_ID, MACRO_ID, OPERATION_ID, CATEGORY_ID, USE_TYPE_ID, WORK_SHIFT_ID, DISTRICT_ID from KPI.KPI_DC_LEVEL 
--
--select STORECODE_ID, TEMPAREA_ID, DC_ID, MACRO_ID, OPERATION_ID, CATEGORY_ID, USE_TYPE_ID, WORK_SHIFT_ID, DISTRICT_ID from KPI.KPI_DC_LEVEL 
--except
--select STORECODE_ID, TEMPAREA_ID, DC_ID, MACRO_ID, OPERATION_ID, CATEGORY_ID, USE_TYPE_ID, WORK_SHIFT_ID, DISTRICT_ID from KPI.v_KPI_DC_LEVEL
*/


CREATE PROCEDURE kpi.SP_SEND_EMAIL_KPI
AS

BEGIN

DROP TABLE IF EXISTS #LIST_BLANK_KPI;
SELECT t.KPI_NAME_ID
     , kpt.PERIOD_TYPE_ID
     , v.STORECODE_ID
     , v.TEMPAREA_ID
     , v.DC_ID
     , v.MACRO_ID
     , v.OPERATION_ID
     , v.CATEGORY_ID
     , v.USE_TYPE_ID
     , v.WORK_SHIFT_ID
     , v.DISTRICT_ID
INTO #LIST_BLANK_KPI
FROM kpi.KPI_USE_TYPE AS t
INNER JOIN [DICT].[KPI_PERIOD_TYPE] (NOLOCK) AS kpt 
	ON kpt.KPI_NAME_ID = t.KPI_NAME_ID
INNER JOIN [DICT].[USE_TYPE] (NOLOCK) AS ut 
	ON ut.ID = t.USE_TYPE_ID 
	AND ut.IS_FACT_MANUAL_FLG = 0
INNER JOIN KPI.V_KPI_DC_LEVEL (NOLOCK) AS v 
	ON v.USE_TYPE_ID = t.USE_TYPE_ID
EXCEPT
SELECT kt.KPI_NAME_ID
     , kt.PERIOD_TYPE_ID
     , kdl.STORECODE_ID
     , kdl.TEMPAREA_ID
     , kdl.DC_ID
     , kdl.MACRO_ID
     , kdl.OPERATION_ID
     , kdl.CATEGORY_ID
     , kdl.USE_TYPE_ID
     , kdl.WORK_SHIFT_ID
     , kdl.DISTRICT_ID
FROM kpi.KPI_TARGET kt
INNER JOIN kpi.KPI_DC_LEVEL AS kdl 
	ON kdl.ID = kt.KPI_DC_LEVEL_ID
WHERE VALID_FROM_DTTM <= getdate() 
	AND VALID_TO_DTTM >= getdate();



DROP TABLE IF EXISTS #ADD_NAME_LIST_BLANK_KPI;
SELECT n.SHORT_NAME_NM          as [Показатель]
     , ut.NAME_NM               as [Тип применимости норматива]
     , gd.NAME_NM               as [Тип периода]
     , d.SHORT_NAME             as [Округ]
     , m.SHORT_NAME_NM          as [Макрорегион]
     , dc.SHORT_NAME_NM         as РЦ
     ---WORK_SHIFT_ID	????
     , sc.SHORT_NAME_NM         as [Лог склад]
     , tm.SHORT_NAME_NM         as [Температурный участок]
     , cat.NAME_NM              as [Товарная категория]
     , o.SHORT_NAME_IG          as [Операция]
INTO #ADD_NAME_LIST_BLANK_KPI
FROM #LIST_BLANK_KPI x    
LEFT JOIN [InfoDB].kpi.KPI_NAME AS n 
	ON n.ID = x.KPI_NAME_ID
LEFT JOIN [InfoDB].KPI.KPI_DC AS dc
	ON dc.ID = x.DC_ID
LEFT JOIN [InfoDB].[DICT].[STORECODE] AS sc
	ON sc.ID = x.STORECODE_ID
LEFT JOIN [DICT].[TEMPAREA] AS tm
	ON tm.ID = x.TEMPAREA_ID
LEFT JOIN [DICT].[MACRO] AS m
	ON m.ID = x.MACRO_ID
LEFT JOIN DBO.tbProizv_Operations AS o
	ON o.CODE = x.OPERATION_ID
LEFT JOIN [DICT].[CATEGORY] AS cat
	ON cat.ID = x.CATEGORY_ID
LEFT JOIN [DICT].[USE_TYPE] AS ut
	ON ut.ID = x.USE_TYPE_ID
LEFT JOIN [DICT].[GENERAL_DICT] AS gd 
	ON gd.ID = x.PERIOD_TYPE_ID
LEFT JOIN [DICT].[DISTRICT] AS d 
	ON d.ID = x.DISTRICT_ID
ORDER BY n.SHORT_NAME_NM, ut.NAME_NM, gd.NAME_NM, d.SHORT_NAME, m.SHORT_NAME_NM, sc.SHORT_NAME_NM;


DECLARE @email_body_data varchar(max);

SET @email_body_data = ''

SELECT @email_body_data = '<table cellpadding=3 cellspacing=0 border=1><tr style="color:White;background-color:SteelBlue;font-weight:bold;"><td>Показатель</td>' +
'<td>Тип применимости норматива</td><td>Тип периода</td><td>Округ</td><td>Макрорегион</td><td>РЦ</td><td>Лог склад</td><td>Температурный участок</td><td>Товарная категория</td><td>Операция</td></tr>' +
CAST ( (
		SELECT /* top 100*/ [Tag] = 1, [Parent] = 0, 
		  [tr!1!td!element] = isnull([Показатель],'')
		, [tr!1!td!element] = isnull([Тип применимости норматива],'')
		, [tr!1!td!element] = isnull([Тип периода],'')
		, [tr!1!td!element] = isnull([Округ],'')
		, [tr!1!td!element] = isnull([Макрорегион],'')
		, [tr!1!td!element] = isnull([РЦ],'')
		, [tr!1!td!element] = isnull([Лог склад],'')
		, [tr!1!td!element] = isnull([Температурный участок],'')
		, [tr!1!td!element] = isnull([Товарная категория],'')
		, [tr!1!td!element] = isnull([Операция],'')
		FROM #ADD_NAME_LIST_BLANK_KPI
		FOR xml explicit
	) AS VARCHAR (MAX) ) + '</table>'

            DECLARE @email		AS	NVARCHAR(MAX) = 'Boris.Khaitov@x5.ru'
            DECLARE @subject    AS	NVARCHAR(MAX) = 'Заполнение нормативов KPI Дашборд РЦ'
            DECLARE @emailBody	AS	NVARCHAR(MAX) =	'Отсутствуют действующие нормативы: ' + @email_body_data
            EXEC msdb.dbo.sp_send_dbmail --@profile_name='SQL Email Profile',
            		@recipients=@email,
            		@subject=@subject,
            		@body=@emailBody,
                    @body_format = 'HTML'


END;

