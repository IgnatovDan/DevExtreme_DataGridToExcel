blog post draft

Начать можно с общего ощущения проблемы в скваде: почти все уже поотвечали на вопросы клиентов "а как мне сделать ХХХ" и у всех ощущение "могло бы быть и лучше"...

В результате появляется список тикетов со схожими задачами (например можно [так оформить эти списки](https://trello.com/c/xX9e0BHE/457-excel-export-hyperlink-as-native-excel-hyperlink)), длина списка сравнивается с длиной списков других схожих задач, что-то выбирается (или не выбирается) для работы на следующий релиз и попадает в roadmap.

Мой список получился вот такой: [DataGrid Export to Excel tickets](https://trello-attachments.s3.amazonaws.com/57ac5e71936273d45a811410/5b910f30c9e013859eeeed2c/x/ff765a41492629f95d9b2d142130505c/Tickets_v4.txt) - тут я сразу еще детальнее сгруппировал тикеты по схожим задачам и составил список более мелких задач тоже в порядке длины списка каждой задачи.

Так как задача с экспортом явно не новая и есть схожие продукты, то я посмотрел как те же самые сценарии решаются "у них" (было больше, не все нашлось): 
- [T686448 - DataGrid.export.customizeExcelCell: Similar topics in other products](https://isc.devexpress.com/Thread/WorkplaceDetails/T686448)
- [JS grids](https://trello-attachments.s3.amazonaws.com/57ac5e71936273d45a811410/5b910f30c9e013859eeeed2c/x/7dbe7113aefb1f96342d6c875223437e/ExportToExcel_Javascript_Grids.txt)
- [DX Controls](https://trello-attachments.s3.amazonaws.com/57ac5e71936273d45a811410/5b910f30c9e013859eeeed2c/x/d61b55785bc9dd532777fdc7928ee284/ExportToExcel_DevExpress_Controls.txt)
- \\\\corp\internal\common\4ignatov\12 

От первого варианта изменений не осталось ничего, решили целиком выкинуть и я сделал новый.

По тестам я делал покрытие по возможности по всем вариантам аргументов, часто перебором. Например, эти тесты проверяют как наш код работает с изменениями строкового значения на другие варианты значений:
- Change string undefined to string
- Change string undefined to number
- Change string null to string
- Change string null to number
- Change string value to undefined
- Change string value to null
- Change string value to string
- Change string value to number
- Change string value to Number.NaN
- Change string value to Number.POSITIVE_INFINITY
- Change string value to Number.NEGATIVE_INFINITY
- Change string value to date
- Change string value to boolean

Это действительно важные сценарии, во многих я сразу получал неадекватные результаты, про некоторые мы получили вопросы клиентов после выпуска. Покрытие тестами по другим публичным интерфейсам и внутренним частям примерно такое же. Всего около 200.

Параллельно с разработкой вели обсуждение с клиентами на гитхабе: [Export to Excel customization improvements](https://github.com/DevExpress/DevExtreme/issues/5165#issuecomment-457442064)

Я сделал технические демки \\corp\internal\common\4ignatov\ExcelExport_TechnicalDemos  и мы их опубликовали как "Live Sandbox":
* [Excel Cell Value](https://codepen.io/dxbykov/pen/MPamQV)
* [Excel Cell Text Appearance](https://codepen.io/dxbykov/pen/ePpWPL)
* [Excel Cell Text Alignment](https://codepen.io/dxbykov/pen/EdGKxK)
* [Excel Cell Background](https://codepen.io/dxbykov/pen/GYpmzj)
* [Excel Cell Number Format](https://codepen.io/dxbykov/pen/ePNLvg)

Несколько раз сделал доку и еще больше раз переделал API: 
- [APIv3](https://trello-attachments.s3.amazonaws.com/57ac5e71936273d45a811410/5b910f30c9e013859eeeed2c/x/09ef4f94596d9db1bcf74aa26dbdb0f7/API_v3.txt)
- [APIv4](https://trello-attachments.s3.amazonaws.com/57ac5e71936273d45a811410/5b910f30c9e013859eeeed2c/x/a67897918ab6a8f6d7fc1a6f83c9098a/API_v4.txt)
- дока: \\\\corp\internal\common\4ignatov\ExcelExport_Docs

Переделал технические демки: \\\\corp\internal\common\4ignatov\ExcelExport_TechnicalDemos_build_2018-10-24_13-17 

Сделал демки на сайт: 
- \\\\corp\internal\common\4ignatov\ExcelExport_Demos
- https://js.devexpress.com/Demos/WidgetsGallery/Demo/DataGrid/ExcelCellCustomization/jQuery/Light/

Сделал заготовку поста: [DevExtreme DataGrid – Excel Data Export Customization Enhancements (v18.2)](https://devexpress.sharepoint.com/sites/devextreme/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2Fdevextreme%2FShared%20Documents%2F18%2E2%2FBlogs%2FExcel%20Data%20Export%20Customization%20Enhancements%2Edocx&parent=%2Fsites%2Fdevextreme%2FShared%20Documents%2F18%2E2%2FBlogs)

Часть сценариев я не успел сделать до кодефриза, их я отложил в столбец "Backlog Export" на доске [devextreme ui widgets backlog](https://trello.com/b/NttE8H8j/devextreme-ui-widgets-backlog)

После выпуска пришли клиенты с вопросами ["а я думал теперь можно ХХХ"](https://trello.com/c/sPryhh4h/105-excel-export-tickets). Для них я сделал изменения сразу, войдут в 18.2.6/18.2.7. Что-то не войдет, про это так и отвечаем "еще нет".

Опубликовали пост [DevExtreme DataGrid – Excel Data Export Customization Enhancements (v18.2)](https://community.devexpress.com/blogs/javascript/archive/2018/11/07/devextreme-datagrid-excel-data-export-customization-enhancements-v18-2.aspx), пообсуждали изменения с клиентами.
