#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область СлужебныеПроцедурыИФункции

// Заполняет список команд печати.
//
// Параметры:
// КомандыПечати – ТаблицаЗначений – состав полей см. в функции УправлениеПечатью.СоздатьКоллекциюКомандПечати.
//
Процедура ДобавитьКомандыПечати(КомандыПечати) Экспорт
	
	
КонецПроцедуры

// СтандартныеПодсистемы.УправлениеДоступом
// См. УправлениеДоступомПереопределяемый.ПриЗаполненииСписковСОграничениемДоступа.
Процедура ПриЗаполненииОграниченияДоступа(Ограничение) Экспорт
	
	// Проверяем подразделение в табличной части, а не в шапке, 
	// так как для поиска значения используется именно подразделение в табличной части.
	// На чтение доступен любой документ, хотя бы одна строка которого разрешена.
	// Но изменить можно только документ, у которого все строки разрешены.

	// От проверки объекта ремонта в табличной части принято решение отказаться по следующим причинам:
	// 1. Техническая: в производительном варианте работы РЛС есть ограничение, можно проверять не больше 
	// 5 полей в табличной части, при этом одна проверка объекта ремонта съедает все эти 5 полей 
	// (флаг "Принадлежит контрагенту", контрагент, организация, подразделение и ссылка), 
	// так что больше ничего добавить уже не выходит.
	// 2. Логическая: Велика вероятность, что конкретные объекты ремонта в данных документах будут указывать редко, 
	// чаще использя более верхнеуровненые - организацию и подразделение. И в конце концов, никто не мешает 
	// при иуказании объекта ремонта дополнительно заполнить еще организацию и подразделение. 
	// Поэтому, можно пренебречь проверкой этого поля в пользу более значимых.

	Ограничение.Текст =
	"РазрешитьЧтение
	|ГДЕ
	|	ЗначениеРазрешено(Организация)
	|	И ДляОднойИзСтрок(ЗначениеРазрешено(ЗначенияПоказателей.Подразделение, NULL КАК ИСТИНА, ПустаяСсылка КАК ИСТИНА)) 
	|;
	|РазрешитьИзменениеЕслиРазрешеноЧтение
	|ГДЕ
	|	ЗначениеРазрешено(Организация)
	|	И ДляВсехСтрок(ЗначениеРазрешено(ЗначенияПоказателей.Подразделение, ПустаяСсылка КАК ИСТИНА))";
	
	Ограничение.ТекстДляВнешнихПользователей = Ограничение.Текст;

КонецПроцедуры
// Конец СтандартныеПодсистемы.УправлениеДоступом

#КонецОбласти

#КонецЕсли