/*
     Функция, при помощи которой мы забираем данные из API Reports Яндекс.Директ/Power BI

     Версия 1.4
     -- добавилась возможность обновления в облаке Power Bi

     Документация по API Reports: https://tech.yandex.ru/direct/doc/reports/reports-docpage/
     Список типов отчетов: https://tech.yandex.ru/direct/doc/reports/type-docpage/
     Список полей для отчета: https://tech.yandex.ru/direct/doc/reports/fields-list-docpage/


     Создатель: Эльдар Забитов (http://zabitov.ru)
*/



let
    pqyd = (Token as text, ClientLogin as nullable text, FieldNames as text, ReportType as text, DateFrom as text, DateTo as text) =>
let
    ClientLogin = if ClientLogin = null then "" else ClientLogin,
// Проверяем на TODAY и YESTERDAY
    DateFrom = Text.Upper(DateFrom),
    DateTo = Text.Upper(DateTo),
    dateFrom = if DateFrom = "TODAY"
        then Date.ToText(DateTime.Date(DateTime.LocalNow()), "yyyy-MM-dd")
    else
        if DateFrom = "YESTERDAY"
            then Date.ToText(Date.AddDays(DateTime.Date(DateTime.LocalNow()), -1), "yyyy-MM-dd")
        else DateFrom,

    dateTo = if DateTo = "TODAY"
        then Date.ToText(DateTime.Date(DateTime.LocalNow()), "yyyy-MM-dd")
    else
        if DateTo = "YESTERDAY"
            then Date.ToText(Date.AddDays(DateTime.Date(DateTime.LocalNow()), -1), "yyyy-MM-dd")
        else DateTo,

// Обрабатываем параметр Fieldnames, делаем список, добавляем нужные значения
// Все это нужно чтобы можно было просто написать список полей через запятую :)
    new = Text.Split(FieldNames, ","),
    ToTable = Table.FromList(new, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    deleteSpace = Table.ReplaceValue(ToTable," ","",Replacer.ReplaceText,{"Column1"}),
    PlusFieldName = Table.AddColumn(deleteSpace, "Custom", each "<FieldNames>"&[Column1]&"</FieldNames>"),
    delete = Table.SelectColumns(PlusFieldName,{"Custom"}),
    transpot = Table.Transpose(delete),
    merge = Table.CombineColumns(transpot, Table.ColumnNames(transpot),Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Merged"),
    fieldnamestext = merge[Merged]{0},
    ReportName = ReportType&"-"&dateFrom&"-"&dateTo&fieldnamestext,

// Присваиваем полученный токен
    AuthKey = "Bearer "&Token,
    url = "https://api.direct.yandex.com/",

// Создаем тело запроса со всеми параметрами
    body =
        "<ReportDefinition xmlns=""http://api.direct.yandex.com/v5/reports"">
        <SelectionCriteria>
        <DateFrom>"&dateFrom&"</DateFrom>
        <DateTo>"&dateTo&"</DateTo>
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
            RelativePath="v5/reports",
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
