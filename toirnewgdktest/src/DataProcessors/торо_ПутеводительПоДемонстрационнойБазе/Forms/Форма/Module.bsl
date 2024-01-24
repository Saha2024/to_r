
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	ЗаполнитьТаблицуРазделов();
	
	ПоказыватьПриОткрытииПрограммы = ОбщегоНазначения.ХранилищеОбщихНастроекЗагрузить("НастройкиТОиР", "ПоказыватьПутеводительПоДемоБазе", Истина);
	
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	
	УстановитьСтартовуюСтраницу();
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ПолеHTMLДокументаПриНажатии(Элемент, ДанныеСобытия, СтандартнаяОбработка)
	
	Если ДанныеСобытия.Свойство("href") И ЗначениеЗаполнено(ДанныеСобытия.href) Тогда
		ИмяОткрываемойСтраницы = СокрЛП(ДанныеСобытия.href);
	Иначе
		Возврат;
	КонецЕсли;
	
	Позиция = Найти(ИмяОткрываемойСтраницы, "ShowMaket=");
	Если Позиция > 0 Тогда
		
		ИмяМакета = Сред(ИмяОткрываемойСтраницы, Позиция + 10);
		МакетИзСписка = СписокМакетов.НайтиСтроки(Новый Структура("URLМакета",ИмяМакета))[0];
		
		Попытка
			
			ПолеHTMLДокумента = МакетИзСписка.Содержание;
			ТекущийМакет = МакетИзСписка.URLМакета;
			
		Исключение
		КонецПопытки;
		
		ДанныеСобытия.Event.returnValue = Ложь;
		Возврат;

	КонецЕсли;	
	
	Позиция = Найти(ИмяОткрываемойСтраницы , "/e1cib/data/");
	Если Позиция > 0 Тогда
		ПоказатьЗначение(,ПерейтиПоНавигационнойСсылкеОбычное(Сред(ИмяОткрываемойСтраницы , Позиция+12)));
		ДанныеСобытия.Event.returnValue = Ложь;
		Возврат;
	Иначе
	  	ОткрытьСправку(ИмяОткрываемойСтраницы);
		СтандартнаяОбработка = Ложь;
		Возврат;
	КонецЕсли;
	
	Если ВРег(Лев(ИмяОткрываемойСтраницы , 4)) = ВРег("http") Тогда
		ПоказатьЗначение(,ПерейтиПоНавигационнойСсылкеОбычное(ИмяОткрываемойСтраницы ));
		ДанныеСобытия.Event.returnValue = Ложь;
		Возврат;
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ПоказыватьПриОткрытииПрограммыПриИзменении(Элемент)
	
	СохранитьНастройкиПоказыватьПриОткрытииПрограммы(ПоказыватьПриОткрытииПрограммы);
	
КонецПроцедуры

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

Функция ПерейтиПоНавигационнойСсылкеОбычное (пСсылка)
	Разделитель = Найти(пСсылка, "__");
	Если Разделитель <> 0 Тогда
		МетаданныеОбъекта = Лев(пСсылка, Найти(пСсылка, "__")-1);		
		Если Найти(МетаданныеОбъекта, "Документ.") <> 0 Тогда
			Менеджер = Документы;	
		ИначеЕсли Найти(МетаданныеОбъекта, "Справочник") <> 0 Тогда
			Менеджер = Справочники;	
		ИначеЕсли Найти(МетаданныеОбъекта, "ПланВидовРасчета") <> 0 Тогда
			Менеджер = ПланыВидовРасчета;	
		ИначеЕсли Найти(МетаданныеОбъекта, "ПланВидовХарактеристик") <> 0 Тогда
			Менеджер = ПланыВидовХарактеристик;	
		ИначеЕсли Найти(МетаданныеОбъекта, "ПланСчетов") <> 0 Тогда
			Менеджер = ПланыСчетов;	
		ИначеЕсли Найти(МетаданныеОбъекта, "БизнесПроцесс") <> 0 Тогда
			Менеджер = БизнесПроцессы;	
		ИначеЕсли Найти(МетаданныеОбъекта, "Задача") <> 0 Тогда
			Менеджер = Задачи;	
		КонецЕсли;	
		Имя = Прав(МетаданныеОбъекта, СтрДлина(МетаданныеОбъекта) - Найти(МетаданныеОбъекта, "."));
		GUIDОбъекта = Новый УникальныйИдентификатор(СтрЗаменить(пСсылка, МетаданныеОбъекта+"__", ""));
		ПолученнаяСсылка = Менеджер[Имя].ПолучитьСсылку(GUIDОбъекта);
		Возврат ПолученнаяСсылка;
		
	КонецЕсли;
КонецФункции

// Функция загружает содержание макетов в таблицу на форму.
// При этом формирует содержание по разделам.
//
// Возвращаемое значение:
//   <ТекстСодержания>   - html макет содержащий раздел "Содержание" полученный из добавленных макетов.
//
&НаСервере
Функция ЗагрузитьМакетыНаФорму()
	
	ТекущийОбъект 		= РеквизитФормыВЗначение("Объект");
	МетаданныеОбработки = ТекущийОбъект.Метаданные(); 
	ТекстСодержания 	= "<H2>Путеводитель по демонстрационной базе</H2>
	                  |";
	
	Для Каждого Макет Из МетаданныеОбработки.Макеты Цикл
		
		МакетHTML 				= ТекущийОбъект.ПолучитьМакет(Макет.Имя);
		URLМакета 				= МакетHTML.ПолучитьДокументHTML().БазовыйURI;
		
		НоваяСтрока 			= СписокМакетов.Добавить();
		НоваяСтрока.Макет		= Макет.Имя;
		НоваяСтрока.URLМакета	= URLМакета;
		НоваяСтрока.Содержание	= МакетHTML.ПолучитьТекст();
		
		ТекстСодержания = ТекстСодержания 
				+ СтрШаблон("<P><A id=%1 href=""ShowMaket=%2"">%3</A></P>"
						,Макет.Имя
						,URLМакета
						,Макет.Синоним);
		
	КонецЦикла;
	
	Возврат ТекстСодержания;

КонецФункции

// Добавляем раздел "Содержание" во все макеты
//
&НаСервере
Процедура УстановитьСодержаниеПоРазделам(ТекстСодержания)

	Для Каждого Макет из СписокМакетов Цикл
		
		Макет.Содержание = СтрЗаменить(Макет.Содержание, 
				"<td id=""contents"" valign=""top"" width=""150""></td>",
				"<td id=""contents"" valign=""top"" width=""150"">" + ТекстСодержания + "</td>");
		
	КонецЦикла;	
			
КонецПроцедуры 

&НаСервере
Процедура ЗаполнитьТаблицуРазделов()

	ТекстСодержания = ЗагрузитьМакетыНаФорму();
	
	УстановитьСодержаниеПоРазделам(ТекстСодержания);

КонецПроцедуры

// Возращает найденную по параметрам строку.
//
// Параметры:
//  <ЗначениеРеквизита>  - <Строка> - что ищем.
//  <ИмяРеквизита>  	 - <Строка> - где ищем.
//
// Возвращаемое значение:
//   <СтрокаТаблицыЗначений>   - строка таблицы значений найденная по значению реквизита. Если не найдено - то Неопределено.
//
&НаКлиенте
Функция НайтиСтрокуМакета(ЗначениеРеквизита, ИмяРеквизита)

	ОтборыДляПоиска = Новый Структура(ИмяРеквизита, ЗначениеРеквизита);
	
	НайденныеСтроки = СписокМакетов.НайтиСтроки(ОтборыДляПоиска);
	
	Если НайденныеСтроки.Количество() = 0 Тогда
	
		Возврат Неопределено;
	
	Иначе
	
		Возврат НайденныеСтроки[0];
	
	КонецЕсли; 
	

КонецФункции 

&НаКлиенте
Процедура УстановитьРаздел(Раздел)

	Если Раздел <> Неопределено Тогда
		
		Попытка
			ПолеHTMLДокумента 	= Раздел.Содержание;
			ТекущийМакет 		= Раздел;
		Исключение
		КонецПопытки;
		
	КонецЕсли; 

КонецПроцедуры 

// Процедура устанавливает стартовую строку макета
//
&НаКлиенте
Процедура УстановитьСтартовуюСтраницу()
	
	РазделНачало = НайтиСтрокуМакета("ОбщееОписание", "Макет");
	
	УстановитьРаздел(РазделНачало);
	
КонецПроцедуры 

&НаСервереБезКонтекста
Процедура СохранитьНастройкиПоказыватьПриОткрытииПрограммы(ПоказыватьПриОткрытииПрограммы)
	
	ОбщегоНазначения.ХранилищеОбщихНастроекСохранить("НастройкиТОиР", "ПоказыватьПутеводительПоДемоБазе", ПоказыватьПриОткрытииПрограммы);
	
КонецПроцедуры

#КонецОбласти
	
