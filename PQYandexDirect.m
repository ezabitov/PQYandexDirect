/*
     Функция, при помощи которой мы забираем данные из API Reports Яндекс.Директ

     Версия 1.1

     Токен можно получить здесь: https://oauth.yandex.ru/authorize?response_type=token&client_id=365a2d0a675c462d90ac145d4f5948cc
     Взято здесь: https://github.com/selesnow/ryandexdirect/blob/master/R/yadirGetToken.R. Автору спасибо
     
     Документация по API Reports: https://tech.yandex.ru/direct/doc/reports/reports-docpage/
     Список типов отчетов: https://tech.yandex.ru/direct/doc/reports/type-docpage/
     Список полей для отчета: https://tech.yandex.ru/direct/doc/reports/fields-list-docpage/


     Создатель: Эльдар Забитов (http://zabitov.ru)
*/

let
    pqyd = (Token as text, ClientLogin as text, ReportName as text, FieldNames as text, ReportType as text, DateFrom as text, DateTo as text) =>
let

// Обрабатываем параметр Fieldnames, делаем список, добавляем нужные значения
// Все это нужно чтобы можно было просто написать список полей через запятую :)
    new = Text.Split(FieldNames, ","),
    ToTable = Table.FromList(new, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    PlusFieldName = Table.AddColumn(ToTable, "Custom", each "<FieldNames>"&[Column1]&"</FieldNames>"),
    delete = Table.SelectColumns(PlusFieldName,{"Custom"}),
    transpot = Table.Transpose(delete),
    merge = Table.CombineColumns(transpot, Table.ColumnNames(transpot),Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Merged"),
    fieldnamestext = merge[Merged]{0},

// Присваиваем полученный токен
    AuthKey = "Bearer "&Token,
    url = "https://api.direct.yandex.com/v5/reports",

// Создаем тело запроса со всеми параметрами
    body = 
        "<ReportDefinition xmlns=""http://api.direct.yandex.com/v5/reports"">
        <SelectionCriteria>
        <DateFrom>"&DateFrom&"</DateFrom>
        <DateTo>"&DateTo&"</DateTo>
        </SelectionCriteria>
        "&fieldnamestext&"
        <ReportName>"&ReportName&"</ReportName>
        <ReportType>"&ReportType&"</ReportType>
        <DateRangeType>CUSTOM_DATE</DateRangeType>
        <Format>TSV</Format>
        <IncludeVAT>YES</IncludeVAT>
        <IncludeDiscount>NO</IncludeDiscount></ReportDefinition>",

// Сам запрос
Source = Web.Contents(url,[
           Content = Text.ToBinary(body) , 

// Заголовки запроса
         Headers = [#"Authorization"=AuthKey ,
                    #"Client-Login"=ClientLogin,
                    #"Accept-Language"="ru",
                    #"Content-Type"="application/x-www-form-urlencoded",
                    #"returnMoneyInMicros" = "false"]
          
             ]   
        )
in
    Source
in
 pqyd