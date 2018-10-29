
/*
     Функция, при помощи которой мы забираем данные из API Reports Яндекс.Директ/Power BI

     Версия 1.6 
     -- Добавлены «Goals» и «AttributionModels»

     Документация по API Reports: https://tech.yandex.ru/direct/doc/reports/reports-docpage/
     Список типов отчетов: https://tech.yandex.ru/direct/doc/reports/type-docpage/
     Список полей для отчета: https://tech.yandex.ru/direct/doc/reports/fields-list-docpage/


     Создатель: Эльдар Забитов (http://zabitov.ru)
*/


let
    pqyd = (Token as text, ClientLogin as nullable text, FieldNames as text, ReportType as text, DateFrom as text, DateTo as text, Goals as nullable text, AttributionModel as nullable text) =>
let
    ClientLogin = if ClientLogin = null then "" else ClientLogin,
    AttributionModel = if List.Contains({"LC", "FC", "LSC"}, Text.Upper(AttributionModel)) = true then "<AttributionModels>"&AttributionModel&"</AttributionModels>" else "",
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

    ReportName = dateFrom&"-"&dateTo&fieldnamestext,

    goals_new = Text.Split(Goals, ","),
    goals_ToTable = Table.FromList(goals_new, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    goals_deleteSpace = Table.ReplaceValue(goals_ToTable," ","",Replacer.ReplaceText,{"Column1"}),
    goals_PlusFieldName = Table.AddColumn(goals_deleteSpace, "Custom", each "<Goals>"&[Column1]&"</Goals>"),
    goals_delete = Table.SelectColumns(goals_PlusFieldName,{"Custom"}),
    goals_transpot = Table.Transpose(goals_delete),
    goals_merge = Table.CombineColumns(goals_transpot, Table.ColumnNames(goals_transpot),Combiner.CombineTextByDelimiter("", QuoteStyle.None),"Merged"),
    goals_text = goals_merge[Merged]{0},


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
        "&goals_text&AttributionModel&fieldnamestext&"
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

           ManualStatusHandling={404, 400},

// Заголовки запроса
         Headers = [#"Authorization"=AuthKey ,
                    #"Client-Login"=ClientLogin,
                    #"Accept-Language"="ru",
                    #"Content-Type"="application/x-www-form-urlencoded",
                    #"returnMoneyInMicros" = "false"]

             ]
        ),
        // Импортируем в Power BI
            importData = Table.FromColumns({Lines.FromBinary(Source,null,null,65001)}),
            removeTopRows = Table.Skip(importData,1),
            removeBottomRows = Table.RemoveLastN(removeTopRows,1),
            split1 = Table.SplitColumn(removeBottomRows, "Column1", Splitter.SplitTextByDelimiter("#(tab)", QuoteStyle.Csv)),
            headers = Table.PromoteHeaders(split1, [PromoteAllScalars=true]),

        // Проверяем, если таблица полученная из Яндекса пустая
            checkIfTableEmpty =

        // Проверяем есть ли финальная таблица
                if Table.IsEmpty(headers)
                    then

        // Проверяем есть ли ответ от сервера с ошибкой
                        if Table.IsEmpty(importData)

        // Если ответа с ошибкой нет - выводим таблицу со строчкой, что ответ таблица еще не готова
                            then #table({"data"}, {{"Ваш отчет еще не готов, попробуйте обновить его позже"}})

        // Если ответ с ошибкой есть - выводим сообщение об ошибке
                            else Xml.Tables(Source,null,65001)[Table]{0}

        // Иначе выводим загруженную таблицу
                    else headers

in
    checkIfTableEmpty
in
 pqyd
