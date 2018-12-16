# Импорт данных Яндекс.Директ в Power BI и Excel

Функция для Power Query и Power BI, которая позволяет забирать статистику из рекламной системы Яндекс.Директ напрямую по API.
Забирать данные будем из [сервиса Reports](https://tech.yandex.ru/direct/doc/reports/reports-docpage/), он предназначен для получения статистики по аккаунту рекламодателя.

[Документация по API Reports](https://tech.yandex.ru/direct/doc/reports/reports-docpage/)

[Список типов отчетов](https://tech.yandex.ru/direct/doc/reports/type-docpage/)

[Список полей для отчета](https://tech.yandex.ru/direct/doc/reports/fields-list-docpage/)

## Как использовать
1. Запускаем расширенный редактор запросов, вставляем код функции.
2. Получаем токен по [ссылке](https://oauth.yandex.ru/authorize?response_type=token&client_id=13f209df5e0b453e86948416a8d90bc9), вставляем
3. Заполняем логин клиента
4. Заполняем логин клиента (Поле ClientLogin).
>Заполняем если агентский аккаунт.
5. Выбираем список полей в отчете.
>Ориентируемся на [список доступных полей](https://tech.yandex.ru/direct/doc/reports/fields-list-docpage/) в документации API Директа. Все поля пишем через запятую, без пробелов.
6. Выбираем тип отчета.
>По [ссылке](https://tech.yandex.ru/direct/doc/reports/type-docpage/) список возможных отчетов. Выбираем нужный.
В примере будем рассматривать «CAMPAIGN_PERFORMANCE_REPORT»
7. Даты начала и конца

>Пишем в формате yyyy-mm-dd. В примере будем рассматривать с 2017-05-01 по 2017-05-10.
Можно использовать "TODAY" и "YESTERDAY".
8. Вызываем функцию.


Подробнее в [блоге](http://zabitov.ru/analitika/yandex-direct-power-query-connector/)
