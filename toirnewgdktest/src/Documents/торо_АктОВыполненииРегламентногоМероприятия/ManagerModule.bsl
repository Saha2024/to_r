#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область СлужебныеПроцедурыИФункции

// Заполняет список команд печати.
//
// Параметры:
// КомандыПечати – ТаблицаЗначений – состав полей см. в функции УправлениеПечатью.СоздатьКоллекциюКомандПечати.
//
Процедура ДобавитьКомандыПечати(КомандыПечати) Экспорт
	
// Акт о выполнении регламентного мероприятия
	КомандаПечати = КомандыПечати.Добавить();
	КомандаПечати.МенеджерПечати = "Документ.торо_АктОВыполненииРегламентногоМероприятия";
	КомандаПечати.Идентификатор = "АктОВыполненииРегламентногоМероприятия";
	КомандаПечати.Представление = НСтр("ru = 'Акт о выполнении регламентного мероприятия'");
	КомандаПечати.Обработчик = "торо_Печать.ЗапроситьУПользователяДополнительныеПараметрыПередПечатью";
	КомандаПечати.СразуНаПринтер = ОбщегоНазначенияВызовСервера.ХранилищеОбщихНастроекЗагрузить(
	"НастройкиТОиР",
	"ПечатьДокументовБезПредварительногоПросмотра",
	Ложь);
	
КонецПроцедуры

// Определяет список команд создания на основании.
//
// Параметры:
//   КомандыСозданияНаОсновании - ТаблицаЗначений - Таблица с командами создания на основании. Для изменения.
//       См. описание 1 параметра процедуры СозданиеНаОснованииПереопределяемый.ПередДобавлениемКомандСозданияНаОсновании().
//   Параметры - Структура - Вспомогательные параметры. Для чтения.
//       См. описание 2 параметра процедуры СозданиеНаОснованииПереопределяемый.ПередДобавлениемКомандСозданияНаОсновании().
//
Процедура ДобавитьКомандыСозданияНаОсновании(КомандыСозданияНаОсновании, Параметры) Экспорт
	
	ИменаОбъектовМетаданных = РегистрыСведений.торо_ВводНаОсновании.ИменаДоступныхОбъектовМетаданныхДляВводаНаОсновании(
		 Метаданные.Документы.торо_АктОВыполненииРегламентногоМероприятия.Имя);
		 
	Для Каждого ИмяОбъектаМетаданных Из ИменаОбъектовМетаданных Цикл
		Документы[ИмяОбъектаМетаданных].ДобавитьКомандуСоздатьНаОсновании(КомандыСозданияНаОсновании);		 
	КонецЦикла;		 
	
	торо_СозданиеНаОсновании.ДобавитьКомандуСоздатьНаОснованииБизнесПроцессЗадание(КомандыСозданияНаОсновании);
	
КонецПроцедуры

// Добавляет команду создания документа "Акт о выполнении регламентного мероприятия".
//
// Параметры:
//   КомандыСозданияНаОсновании - ТаблицаЗначений - Таблица с командами создания на основании. Для изменения.
//       См. описание 1 параметра процедуры СозданиеНаОснованииПереопределяемый.ПередДобавлениемКомандСозданияНаОсновании().
//
Функция ДобавитьКомандуСоздатьНаОсновании(КомандыСозданияНаОсновании) Экспорт
	
	Если ПравоДоступа("Добавление", Метаданные.Документы.торо_АктОВыполненииРегламентногоМероприятия) Тогда
        КомандаСоздатьНаОсновании = КомандыСозданияНаОсновании.Добавить();
        КомандаСоздатьНаОсновании.Менеджер = Метаданные.Документы.торо_АктОВыполненииРегламентногоМероприятия.ПолноеИмя();
        КомандаСоздатьНаОсновании.Представление = ОбщегоНазначения.ПредставлениеОбъекта(Метаданные.Документы.торо_АктОВыполненииРегламентногоМероприятия);
        КомандаСоздатьНаОсновании.РежимЗаписи = "Проводить";
		КомандаСоздатьНаОсновании.ФункциональныеОпции = "торо_ИспользоватьРегламентныеМероприятия";
        Возврат КомандаСоздатьНаОсновании;
	КонецЕсли; 
	
    Возврат Неопределено;
	
КонецФункции

// Сформировать печатные формы объектов.
//
// ВХОДЯЩИЕ:
//   ИменаМакетов    - Строка    - Имена макетов, перечисленные через запятую.
//   МассивОбъектов  - Массив    - Массив ссылок на объекты которые нужно распечатать.
//   ПараметрыПечати - Структура - Структура дополнительных параметров печати.
//
// ИСХОДЯЩИЕ:
//   КоллекцияПечатныхФорм - Таблица значений - Сформированные табличные документы.
//   ПараметрыВывода       - Структура        - Параметры сформированных табличных документов.
//
Процедура Печать(МассивОбъектов, ПараметрыПечати, КоллекцияПечатныхФорм, ОбъектыПечати, ПараметрыВывода) Экспорт
	
	Если УправлениеПечатью.НужноПечататьМакет(КоллекцияПечатныхФорм, "АктОВыполненииРегламентногоМероприятия") Тогда
		УправлениеПечатью.ВывестиТабличныйДокументВКоллекцию(КоллекцияПечатныхФорм, 
		"АктОВыполненииРегламентногоМероприятия", 
		"Акт о выполенении регламентного мероприятия", 
		ПечатьАктОВыполненииРегламентногоМероприятия(МассивОбъектов, ПараметрыПечати),
		,
		"Документ.торо_АктОВыполненииРегламентногоМероприятия.ПФ_MXL_АктОВыполненииРегламентногоМероприятия");		
	КонецЕсли;
		
КонецПроцедуры

Процедура ПриОпределенииНастроекВерсионированияОбъектов(Настройки) Экспорт
	
	Настройки.ПриПолученииСлужебныхРеквизитов = Истина;
	
КонецПроцедуры

Процедура ПриПолученииСлужебныхРеквизитов(Реквизиты) Экспорт
	
	Реквизиты.Добавить("СтатусДокумента");
	Реквизиты.Добавить("ИсторияСтатусов.*");
	Реквизиты.Добавить("Согласующие.*");
		
КонецПроцедуры

Функция ПечатьАктОВыполненииРегламентногоМероприятия(МассивОбъектов, ПараметрыПечати)
	
	ТабДок = Новый ТабличныйДокумент;
	Макет = УправлениеПечатью.МакетПечатнойФормы("Документ.торо_АктОВыполненииРегламентногоМероприятия.ПФ_MXL_АктОВыполненииРегламентногоМероприятия");
	ИспользоватьКомиссии = Константы.торо_ИспользоватьКомиссии.Получить();
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
	"ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	торо_АктОВыполненииРегламентногоМероприятияМероприятия.Ссылка КАК Документ,
	|	торо_АктОВыполненииРегламентногоМероприятияМероприятия.ОбъектРемонта КАК ОбъектРемонта,
	|	торо_АктОВыполненииРегламентногоМероприятияМероприятия.ВидРемонтныхРабот КАК ВидРемонтныхРабот,
	|	торо_АктОВыполненииРегламентногоМероприятияМероприятия.ДокументИсточник КАК ДокументИсточник,
	|	торо_АктОВыполненииРегламентногоМероприятияМероприятия.ДатаНачала КАК ДатаНачала,
	|	торо_АктОВыполненииРегламентногоМероприятияМероприятия.ДатаОкончания КАК ДатаОкончания,
	|	торо_АктОВыполненииРегламентногоМероприятияМероприятия.ID КАК ID,
	|	торо_АктОВыполненииРегламентногоМероприятия.Номер КАК Номер,
	|	торо_АктОВыполненииРегламентногоМероприятия.Дата КАК Дата,
	|	торо_АктОВыполненииРегламентногоМероприятия.Организация КАК Организация,
	|	торо_АктОВыполненииРегламентногоМероприятия.Подразделение КАК Подразделение";
	Если ИспользоватьКомиссии Тогда
		Запрос.Текст = Запрос.Текст + 
		",
		|	торо_АктОВыполненииРегламентногоМероприятия.Комиссия КАК Комиссия,
		|	торо_АктОВыполненииРегламентногоМероприятия.Комиссия.Председатель КАК ПредседательКомиссии";		
	КонецЕсли;	
	Запрос.Текст = Запрос.Текст + 
	" ИЗ
	|	Документ.торо_АктОВыполненииРегламентногоМероприятия КАК торо_АктОВыполненииРегламентногоМероприятия
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.торо_АктОВыполненииРегламентногоМероприятия.Мероприятия КАК торо_АктОВыполненииРегламентногоМероприятияМероприятия
	|		ПО торо_АктОВыполненииРегламентногоМероприятия.Ссылка = торо_АктОВыполненииРегламентногоМероприятияМероприятия.Ссылка
	|ГДЕ
	|	торо_АктОВыполненииРегламентногоМероприятия.Ссылка В(&Ссылка)
	|	И торо_АктОВыполненииРегламентногоМероприятияМероприятия.ID В(&МассивID)
	|ИТОГИ
	|	МИНИМУМ(ДатаНачала),
	|	МАКСИМУМ(ДатаОкончания),
	|	МАКСИМУМ(Номер),
	|	МАКСИМУМ(Дата),
	|	МАКСИМУМ(Организация),
	|	МАКСИМУМ(Подразделение)
	|ПО
	|	Документ";
	
	Запрос.УстановитьПараметр("Ссылка", МассивОбъектов);
	Если ПараметрыПечати.Свойство("МассивID") Тогда
		Запрос.УстановитьПараметр("МассивID", ПараметрыПечати.МассивID);
	Иначе
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "И торо_АктОВыполненииРегламентногоМероприятияМероприятия.ID В(&МассивID)", "");
	КонецЕсли;
	
	ВыборкаДокументов = Запрос.Выполнить().Выбрать(ОбходРезультатаЗапроса.ПоГруппировкам, "Документ");
	ЭтоПервый = Истина;
	
	Пока ВыборкаДокументов.Следующий() Цикл
		Документ = ВыборкаДокументов.Документ;
		ДатаНачала = ВыборкаДокументов.ДатаНачала;
		ВыборкаМероприятий = ВыборкаДокументов.Выбрать(ОбходРезультатаЗапроса.Прямой);
		
		Запрос = Новый Запрос;
		Запрос.Текст = 
		"ВЫБРАТЬ
		|	ВЫБОР
		|		КОГДА торо_АктРемонтныеРаботы.РемонтнаяРабота ССЫЛКА Справочник.торо_ТехнологическиеОперации
		|			ТОГДА ВЫБОР
		|					КОГДА торо_АктРемонтныеРаботы.РемонтнаяРабота.НаименованиеПолное = """"
		|						ТОГДА ТехнологическиеОперации.Наименование
		|					ИНАЧЕ торо_АктРемонтныеРаботы.РемонтнаяРабота.НаименованиеПолное
		|				  КОНЕЦ
		|		ИНАЧЕ ПРЕДСТАВЛЕНИЕССЫЛКИ(торо_ВерсииТехКартСрезПоследних.ТехКарта)
		|	КОНЕЦ КАК Работа,
		|	торо_АктРемонтныеРаботы.ID КАК ID,
		|	торо_АктРемонтныеРаботы.Родитель_ID КАК Родитель_ID,
		|	торо_АктРемонтныеРаботы.РемонтыОборудования_ID КАК РемонтыОборудования_ID,
		|	ВЫБОР
		|		КОГДА торо_АктРемонтныеРаботы.РемонтнаяРабота ССЫЛКА Справочник.торо_ТехнологическиеОперации
		|			ТОГДА ТехнологическиеОперации.Комментарий
		|		ИНАЧЕ торо_ВерсииТехКартСрезПоследних.ТехКарта.Комментарий
		|	КОНЕЦ КАК Комментарий,
		|	торо_АктРемонтныеРаботы.Количество КАК Количество,
		|	ТехнологическиеОперации.БазоваяЕдиницаИзмерения КАК ЕдиницаИзмерения
		|ИЗ
		|	Документ.торо_АктОВыполненииРегламентногоМероприятия.РемонтныеРаботы КАК торо_АктРемонтныеРаботы
		|		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.торо_ТехнологическиеОперации КАК ТехнологическиеОперации
		|		ПО торо_АктРемонтныеРаботы.РемонтнаяРабота = ТехнологическиеОперации.Ссылка
		|		ЛЕВОЕ СОЕДИНЕНИЕ РегистрСведений.торо_ВерсииТехКарт.СрезПоследних(&ДатаНачала, ) КАК торо_ВерсииТехКартСрезПоследних
		|		ПО торо_АктРемонтныеРаботы.РемонтнаяРабота = торо_ВерсииТехКартСрезПоследних.ИдентификаторТехКарты
		|ГДЕ
		|	торо_АктРемонтныеРаботы.Ссылка = &Акт";
		
		Запрос.УстановитьПараметр("Акт", Документ);	
		Запрос.УстановитьПараметр("ДатаНачала", ДатаНачала);
		тзСписокРаботИсх = Запрос.Выполнить().Выгрузить();
		
		Пока ВыборкаМероприятий.Следующий() Цикл
			Мероприятие  = ВыборкаМероприятий;
			ОбщийИтог = 0;
			
			Если НЕ ЭтоПервый Тогда
				ТабДок.ВывестиГоризонтальныйРазделительСтраниц();
			КонецЕсли;
			ЭтоПервый = Ложь;
			
			// Шапка
			Область = Макет.ПолучитьОбласть("Шапка");
			Область.Параметры.НомерДокумента = ВыборкаДокументов.Номер;
			Область.Параметры.Дата	= Формат(ВыборкаДокументов.Дата,"ДЛФ=Д");
			Область.Параметры.ДатаКонец = Формат(Мероприятие.ДатаОкончания,"ДЛФ=Д");
			Область.Параметры.ДатаНач = Формат(Мероприятие.ДатаНачала,"ДЛФ=Д");
			
			ТабДок.Вывести(Область);
			
			// Заголовок
			Область = Макет.ПолучитьОбласть("Заголовок");
			Область.Параметры.ПодразделениеИсполнитель = ВыборкаДокументов.Подразделение;
						
			ТабДок.Вывести(Область);
			
			// Комиссия			
			Если ИспользоватьКомиссии Тогда
				ЗапросКомиссии = Новый Запрос;
				ЗапросКомиссии.УстановитьПараметр("Комиссия", ВыборкаМероприятий.Комиссия);
				ЗапросКомиссии.Текст = "ВЫБРАТЬ
				                      |	торо_КомиссииСоставКомиссии.ЧленКомиссии КАК ЧленКомиссии
				                      |ИЗ
				                      |	Справочник.торо_Комиссии.СоставКомиссии КАК торо_КомиссииСоставКомиссии
				                      |ГДЕ
				                      |	торо_КомиссииСоставКомиссии.Ссылка = &Комиссия";
				ТЗЧленыКомиссии = ЗапросКомиссии.Выполнить().Выгрузить();
				
				Если ЗначениеЗаполнено(ВыборкаМероприятий.ПредседательКомиссии) Тогда 
					ОбластьПредседатель = Макет.ПолучитьОбласть("ПредседательКомиссии");
					ОбластьПредседатель.Параметры.ПредседательКомиссии = ВыборкаМероприятий.ПредседательКомиссии;
					ТабДок.Вывести(ОбластьПредседатель);
				КонецЕсли;
				
				Если ТЗЧленыКомиссии.Количество() Тогда
					ПерваяСтрока = Истина;
					ОбластьЧленКомиссииСЗаголовком = Макет.ПолучитьОбласть("ЧленКомиссииСЗаголовком");
					
					Для Каждого Строка Из ТЗЧленыКомиссии Цикл
						Если ПерваяСтрока Тогда
							ПерваяСтрока = Ложь;
							ОбластьЧленКомиссииСЗаголовком.Параметры.ЧленКомиссии = Строка.ЧленКомиссии;
							ТабДок.Вывести(ОбластьЧленКомиссииСЗаголовком);
							Продолжить;
						КонецЕсли;
						
						ОбластьЧленКомиссии = Макет.ПолучитьОбласть("ЧленКомиссии");
						ОбластьЧленКомиссии.Параметры.ЧленКомиссии = Строка.ЧленКомиссии;
						ТабДок.Вывести(ОбластьЧленКомиссии);
					КонецЦикла;
					
					Если ПерваяСтрока Тогда
						ТабДок.Вывести(ОбластьЧленКомиссииСЗаголовком);
					КонецЕсли;
				КонецЕсли;
			КонецЕсли;
			
			// Шапка списка объектов ремонта
			Область = Макет.ПолучитьОбласть("ШапкаТЧСписокОР");
			ТабДок.Вывести(Область);			
			
			// Список объектов ремонта
			Область = Макет.ПолучитьОбласть("СтрокаТЧСписокОР");
			Запрос = Новый Запрос;
			Запрос.Текст =
				"ВЫБРАТЬ РАЗРЕШЕННЫЕ
				|	торо_АктОВыполненииРегламентногоМероприятияМероприятияОбъектов.ОбъектРемонта КАК ОбъектРемонта,
				|	торо_АктОВыполненииРегламентногоМероприятияМероприятияОбъектов.ОбъектРемонта.АдресМестонахождения КАК Местонахождение
				|ИЗ
				|	Документ.торо_АктОВыполненииРегламентногоМероприятия.МероприятияОбъектов КАК торо_АктОВыполненииРегламентногоМероприятияМероприятияОбъектов
				|ГДЕ
				|	торо_АктОВыполненииРегламентногоМероприятияМероприятияОбъектов.Ссылка = &Ссылка
				|	И торо_АктОВыполненииРегламентногоМероприятияМероприятияОбъектов.Мероприятие_ID = &ID";
			Запрос.УстановитьПараметр("Ссылка", Документ);
			Запрос.УстановитьПараметр("ID", Мероприятие.ID);
			
			Результат = Запрос.Выполнить();
			Выборка = Результат.Выбрать();
			
			КоличествоОбъектовРемонтаПоМероприятию = 0;
			Пока Выборка.Следующий() Цикл 
				КоличествоОбъектовРемонтаПоМероприятию = КоличествоОбъектовРемонтаПоМероприятию + 1;
				Область.Параметры.ОбъектРемонта  = торо_ЗаполнениеДокументов.ПолучитьПредоставленияОРДляПечати(Выборка.ОбъектРемонта);
				Область.Параметры.Местонахождение = Выборка.Местонахождение;
				Область.Параметры.Номер = КоличествоОбъектовРемонтаПоМероприятию;
				ТабДок.Вывести(Область);		
			КонецЦикла;

			
			// ВыполненныеРаботыШапка
			Область = Макет.ПолучитьОбласть("ВыполненныеРаботыШапка");
			ТабДок.Вывести(Область);
			
			// ВыполненныеРаботы
			Область = Макет.ПолучитьОбласть("ВыполненныеРаботы");
			
			// Накладываем отбор по Мероприятию. Если делать это в запросе - теряется последовательность мероприятий.
			ПараметрыОтбора = Новый Структура("РемонтыОборудования_ID",Мероприятие.ID);
			тзСписокРабот = тзСписокРаботИсх.Скопировать(ПараметрыОтбора);	
			Ном = "1";
			МассивID = Новый Массив;
			МассивСтрок = тзСписокРабот.НайтиСтроки(Новый Структура("Родитель_ID",""));
			
			// Определяем используются ли комментарии или нет
			ВывестиМероприятия(МассивСтрок,тзСписокРабот,Область,ТабДок,МассивID,Ном,Истина,Истина);
			
			
			// МатериалыШапка
			ИспользоватьХарактеристики = Константы.торо_ИспользоватьХарактеристикиНоменклатуры.Получить(); 
			Если ИспользоватьХарактеристики Тогда  
				ОбластьШапка = Макет.ПолучитьОбласть("МатериалыФОШапка");
				Область      = Макет.ПолучитьОбласть("МатериалыФО");
			Иначе 
				ОбластьШапка = Макет.ПолучитьОбласть("МатериалыШапка");
				Область      = Макет.ПолучитьОбласть("Материалы");
			КонецЕсли;
			ТабДок.Вывести(ОбластьШапка);
			
			// Материалы
			Запрос = Новый Запрос;
			Запрос.Текст = "ВЫБРАТЬ
			               |	СУММА(МатериальныеЗатраты.Количество) КАК Количество,
			               |	СУММА(МатериальныеЗатраты.Сумма) КАК Сумма,
			               |	МатериальныеЗатраты.Номенклатура КАК Номенклатура,
			               |	МатериальныеЗатраты.ХарактеристикаНоменклатуры КАК ХарактеристикаНоменклатуры,
			               |	МатериальныеЗатраты.ЕдиницаИзмерения КАК ЕдиницаИзмерения
			               |ПОМЕСТИТЬ МатЗатраты_Упаковки
			               |ИЗ
			               |	Документ.торо_АктОВыполненииРегламентногоМероприятия.МатериальныеЗатраты КАК МатериальныеЗатраты
			               |ГДЕ
			               |	МатериальныеЗатраты.Ссылка = &Ссылка
			               |	И МатериальныеЗатраты.РемонтыОборудования_ID = &ID
			               |
			               |СГРУППИРОВАТЬ ПО
			               |	МатериальныеЗатраты.Номенклатура,
			               |	МатериальныеЗатраты.ХарактеристикаНоменклатуры,
			               |	МатериальныеЗатраты.ЕдиницаИзмерения
			               |
			               |ИНДЕКСИРОВАТЬ ПО
			               |	ЕдиницаИзмерения
			               |;
			               |
			               |////////////////////////////////////////////////////////////////////////////////
			               |ВЫБРАТЬ
			               |	МатериальныеЗатраты.Количество * ЕСТЬNULL(спрУпаковки.Коэффициент, 1) КАК Количество,
			               |	МатериальныеЗатраты.Сумма КАК Сумма,
			               |	МатериальныеЗатраты.Номенклатура КАК Номенклатура,
			               |	МатериальныеЗатраты.ХарактеристикаНоменклатуры КАК ХарактеристикаНоменклатуры,
			               |	спрНоменклатура.ЕдиницаИзмерения КАК ЕдиницаИзмерения
			               |ПОМЕСТИТЬ МатЗатраты
			               |ИЗ
			               |	МатЗатраты_Упаковки КАК МатериальныеЗатраты
			               |		ЛЕВОЕ СОЕДИНЕНИЕ Справочник.УпаковкиНоменклатуры КАК спрУпаковки
			               |		ПО МатериальныеЗатраты.ЕдиницаИзмерения = спрУпаковки.Ссылка
			               |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.Номенклатура КАК спрНоменклатура
			               |		ПО МатериальныеЗатраты.Номенклатура = спрНоменклатура.Ссылка
			               |
			               |ИНДЕКСИРОВАТЬ ПО
			               |	Номенклатура,
			               |	ХарактеристикаНоменклатуры,
			               |	ЕдиницаИзмерения
			               |;
			               |
			               |////////////////////////////////////////////////////////////////////////////////
			               |ВЫБРАТЬ
			               |	ЦеныНоменклатурыСрезПоследних.ВидЦены КАК ВидЦены,
			               |	ЦеныНоменклатурыСрезПоследних.Цена КАК Цена,
			               |	ЦеныНоменклатурыСрезПоследних.Валюта КАК Валюта,
			               |	ЦеныНоменклатурыСрезПоследних.ЕдиницаИзмерения КАК ЕдиницаИзмерения,
			               |	ЦеныНоменклатурыСрезПоследних.Номенклатура КАК Номенклатура,
			               |	ЦеныНоменклатурыСрезПоследних.Характеристика КАК Характеристика
			               |ПОМЕСТИТЬ ЦеныНоменклатурыСрезПоследних
			               |ИЗ
			               |	РегистрСведений.ЦеныНоменклатуры.СрезПоследних(
			               |			&Дата,
			               |			ВидЦены = &ТипЦены
			               |				И Валюта = &Валюта
			               |				И (Номенклатура, Характеристика, ЕдиницаИзмерения) В
			               |					(ВЫБРАТЬ
			               |						Таб.Номенклатура,
			               |						Таб.ХарактеристикаНоменклатуры,
			               |						Таб.ЕдиницаИзмерения
			               |					ИЗ
			               |						МатЗатраты КАК Таб)) КАК ЦеныНоменклатурыСрезПоследних
			               |
			               |ИНДЕКСИРОВАТЬ ПО
			               |	Номенклатура,
			               |	Характеристика,
			               |	ЕдиницаИзмерения
			               |;
			               |
			               |////////////////////////////////////////////////////////////////////////////////
			               |ВЫБРАТЬ
			               |	торо_АктОВыполненииРегламентногоМероприятияМатериальныеЗатраты.Количество КАК Количество,
			               |	торо_АктОВыполненииРегламентногоМероприятияМатериальныеЗатраты.Сумма КАК Сумма,
			               |	торо_АктОВыполненииРегламентногоМероприятияМатериальныеЗатраты.Номенклатура КАК Номенклатура,
			               |	торо_АктОВыполненииРегламентногоМероприятияМатериальныеЗатраты.ХарактеристикаНоменклатуры КАК ХарактеристикаНоменклатуры,
			               |	ЕСТЬNULL(ЦеныНоменклатурыСрезПоследних.ВидЦены, ЗНАЧЕНИЕ(Справочник.ВидыЦен.ПустаяСсылка)) КАК ТипЦены,
			               |	ЕСТЬNULL(ЦеныНоменклатурыСрезПоследних.Цена, 0) КАК Цена,
			               |	ЕСТЬNULL(ЦеныНоменклатурыСрезПоследних.Валюта, ЗНАЧЕНИЕ(Справочник.Валюты.ПустаяСсылка)) КАК Валюта,
			               |	торо_АктОВыполненииРегламентногоМероприятияМатериальныеЗатраты.Номенклатура.НаименованиеПолное КАК НоменклатураНаименованиеПолное,
			               |	торо_АктОВыполненииРегламентногоМероприятияМатериальныеЗатраты.Номенклатура.Код КАК НоменклатураКод,
			               |	торо_АктОВыполненииРегламентногоМероприятияМатериальныеЗатраты.Номенклатура.ЕдиницаИзмерения КАК НоменклатураЕдиницаИзмерения
			               |ИЗ
			               |	МатЗатраты КАК торо_АктОВыполненииРегламентногоМероприятияМатериальныеЗатраты
			               |		ЛЕВОЕ СОЕДИНЕНИЕ ЦеныНоменклатурыСрезПоследних КАК ЦеныНоменклатурыСрезПоследних
			               |		ПО торо_АктОВыполненииРегламентногоМероприятияМатериальныеЗатраты.Номенклатура = ЦеныНоменклатурыСрезПоследних.Номенклатура
			               |			И торо_АктОВыполненииРегламентногоМероприятияМатериальныеЗатраты.ХарактеристикаНоменклатуры = ЦеныНоменклатурыСрезПоследних.Характеристика
			               |			И торо_АктОВыполненииРегламентногоМероприятияМатериальныеЗатраты.ЕдиницаИзмерения = ЦеныНоменклатурыСрезПоследних.ЕдиницаИзмерения";
			
			Запрос.УстановитьПараметр("Ссылка", Документ);
			Запрос.УстановитьПараметр("Дата", Мероприятие.ДатаНачала);
			Запрос.УстановитьПараметр("ТипЦены",Константы.торо_ТипЦеныДляРасчетаФактическойСебестоимостиРемонта.Получить());
			Запрос.УстановитьПараметр("Валюта",Константы.ВалютаУправленческогоУчета.Получить());
			Запрос.УстановитьПараметр("ID", Мероприятие.ID);
			
			Результат = Запрос.Выполнить();
			Выборка = Результат.Выбрать();
			
			счетчик = 1;
			Итого=0;
			Пока Выборка.Следующий() Цикл
				
				Область.Параметры.Материал   = Выборка.НоменклатураНаименованиеПолное;
				Если ИспользоватьХарактеристики Тогда  
					Область.Параметры.Характеристика = Строка(Выборка.ХарактеристикаНоменклатуры);
				КонецЕсли;
				Область.Параметры.Ед 		 = Выборка.НоменклатураЕдиницаИзмерения;
				Область.Параметры.Код 		 = Выборка.НоменклатураКод;
				Область.Параметры.Количество = Выборка.Количество;
				Если ЗначениеЗаполнено(Выборка.Цена) Тогда
					Область.Параметры.Стоимость  = Выборка.Цена*Выборка.Количество;
				Иначе
					Область.Параметры.Стоимость  = 0;		
				КонецЕсли; 
				
				Итого = Итого +  Область.Параметры.Стоимость;
				Область.Параметры.Номер      = счетчик;
				счетчик = счетчик + 1;
				ТабДок.Вывести(Область);		
				
			КонецЦикла;
						
			Область = Макет.ПолучитьОбласть("ИтогоМатериалы");
			Область.Параметры.Итого = Итого;
			ОбщийИтог = ОбщийИтог + Итого;
			ТабДок.Вывести(Область);
			
				
			Область = Макет.ПолучитьОбласть("СобственныеШапка");
			ТабДок.Вывести(Область);
			
			Область = Макет.ПолучитьОбласть("Собственные");
			Запрос = Новый Запрос;
			Запрос.Текст = "ВЫБРАТЬ
			|	торо_АктОВыполненииРегламентногоМероприятияТрудовыеЗатраты.Квалификация КАК Квалификация,
			|	СУММА(торо_АктОВыполненииРегламентногоМероприятияТрудовыеЗатраты.ВремяРаботы) КАК ВремяРаботы,
			|	СУММА(торо_АктОВыполненииРегламентногоМероприятияТрудовыеЗатраты.СуммаКНачислению) КАК СуммаКНачислению,
			|	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ торо_АктОВыполненииРегламентногоМероприятияТрудовыеЗатраты.ID) КАК ID,
			|	торо_АктОВыполненииРегламентногоМероприятияТрудовыеЗатраты.Сотрудник КАК Сотрудник
			|ИЗ
			|	Документ.торо_АктОВыполненииРегламентногоМероприятия.ТрудовыеЗатраты КАК торо_АктОВыполненииРегламентногоМероприятияТрудовыеЗатраты
			|ГДЕ
			|	торо_АктОВыполненииРегламентногоМероприятияТрудовыеЗатраты.Ссылка = &Ссылка
			|	И торо_АктОВыполненииРегламентногоМероприятияТрудовыеЗатраты.РемонтыОборудования_ID = &ID
			|
			|СГРУППИРОВАТЬ ПО
			|	торо_АктОВыполненииРегламентногоМероприятияТрудовыеЗатраты.Квалификация,
			|	торо_АктОВыполненииРегламентногоМероприятияТрудовыеЗатраты.Сотрудник";
			
			
			Запрос.УстановитьПараметр("Ссылка", Документ);
			Запрос.УстановитьПараметр("ID", Мероприятие.ID);
			
			Результат = Запрос.Выполнить();
			Выборка = Результат.Выбрать();
			
			счетчик = 1;
			Итого = 0;
			Пока Выборка.Следующий() Цикл
				
				Область.Параметры.Квалификация = Выборка.Квалификация;
				Область.Параметры.ФИО = Выборка.Сотрудник;  
				ВремяРаботы = Выборка.ВремяРаботы * КоличествоОбъектовРемонтаПоМероприятию;
				Область.Параметры.ВремяРаботы  = торо_ОбщегоНазначенияКлиентСервер.СформироватьЗаголовокПоПродолжительности(ВремяРаботы);
				Область.Параметры.Стоимость    = Выборка.СуммаКНачислению * КоличествоОбъектовРемонтаПоМероприятию;
				Область.Параметры.Номер        = счетчик;
				
				Итого = Итого + Область.Параметры.Стоимость;
				счетчик = счетчик + 1;
				ТабДок.Вывести(Область);		
				
			КонецЦикла;
			
			Область = Макет.ПолучитьОбласть("ИтогоСобственные");
			Область.Параметры.Итого = Итого;
			ОбщийИтог = ОбщийИтог + Итого;
			ТабДок.Вывести(Область);
			// ИтогРаботы
				
			
			// Подрядчики
				
			Область = Макет.ПолучитьОбласть("ПодрядчикШапка");
			ТабДок.Вывести(Область);
			
			Область = Макет.ПолучитьОбласть("Подрядчик");
			Запрос = Новый Запрос;
			Запрос.Текст = "ВЫБРАТЬ
			|	АктВыполненныеРаботы.РемонтнаяРабота КАК РемонтнаяРабота,
			|	АктВыполненныеРаботы.ID КАК ID
			|ПОМЕСТИТЬ АктВыполненныеРаботы
			|ИЗ
			|	&АктВыполненныеРаботы КАК АктВыполненныеРаботы
			|ГДЕ
			|	НЕ АктВыполненныеРаботы.РемонтнаяРабота ССЫЛКА Справочник.торо_ТехнологическиеОперации
			|;
			|
			|////////////////////////////////////////////////////////////////////////////////
			|ВЫБРАТЬ
			|	ПодрядчикиРемонтныхРабот.Контрагент КАК Контрагент,
			|	ПодрядчикиРемонтныхРабот.ID КАК ID,
			|	ПодрядчикиРемонтныхРабот.Сумма КАК Сумма
			|ПОМЕСТИТЬ ПодрядчикиРемонтныхРабот
			|ИЗ
			|	&ПодрядчикиРемонтныхРабот КАК ПодрядчикиРемонтныхРабот
			|;
			|
			|////////////////////////////////////////////////////////////////////////////////
			|ВЫБРАТЬ
			|	торо_АктОВыполненииРегламентногоМероприятияПодрядчикиРемонтныхРабот.Контрагент КАК Контрагент,
			|	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ торо_АктОВыполненииРегламентногоМероприятияПодрядчикиРемонтныхРабот.ID) КАК ID,
			|	СУММА(торо_АктОВыполненииРегламентногоМероприятияПодрядчикиРемонтныхРабот.Сумма) КАК Сумма
			|ИЗ
			|	ПодрядчикиРемонтныхРабот КАК торо_АктОВыполненииРегламентногоМероприятияПодрядчикиРемонтныхРабот
			|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ АктВыполненныеРаботы КАК АктВыполненныеРаботы
			|		ПО торо_АктОВыполненииРегламентногоМероприятияПодрядчикиРемонтныхРабот.ID = АктВыполненныеРаботы.ID
			|
			|СГРУППИРОВАТЬ ПО
			|	торо_АктОВыполненииРегламентногоМероприятияПодрядчикиРемонтныхРабот.Контрагент";
			
			Запрос.УстановитьПараметр("АктВыполненныеРаботы",Документ.РемонтныеРаботы.Выгрузить(Новый Структура("РемонтыОборудования_ID",Мероприятие.ID),"РемонтнаяРабота, ID"));
			Запрос.УстановитьПараметр("ПодрядчикиРемонтныхРабот",Документ.ПодрядчикиРемонтныхРабот.Выгрузить());
			
			Результат = Запрос.Выполнить();
			Выборка = Результат.Выбрать();
			счетчик = 1;
			Итого = 0;
			Пока Выборка.Следующий() Цикл
				
				Область.Параметры.Подрядчик   = Выборка.Контрагент;
				Область.Параметры.Стоимость  = Выборка.Сумма;
				Итого = Итого +  Область.Параметры.Стоимость;
				Область.Параметры.Номер      = счетчик;
				счетчик = счетчик + 1;
				ТабДок.Вывести(Область);
				
			КонецЦикла; 
			
			Область = Макет.ПолучитьОбласть("ИтогоПодрядчик");
			Область.Параметры.Итого = Итого;
			ОбщийИтог = ОбщийИтог + Итого;
			ТабДок.Вывести(Область);
			// ИтогРаботы
				
			
			// Подвал
			Область = Макет.ПолучитьОбласть("Подвал");
			Область.Параметры.СуммаУслуг = ОбщийИтог; 
			ФормСтрока = "Л = ru_RU; ДП = Истина";
			ПарамПред  = Справочники.Валюты.НайтиПоНаименованию(Константы.ВалютаУправленческогоУчета.Получить()).ПараметрыПрописи;
			СуммаУслуг = " " + Строка(Константы.ВалютаУправленческогоУчета.Получить())+" ("+ ЧислоПрописью(ОбщийИтог, ФормСтрока,ПарамПред)+")";
			Область.Параметры.СуммаУслугПрописью = СуммаУслуг;
			
			ТабДок.Вывести(Область);
			
			// Подписи комиссии			
			Если ИспользоватьКомиссии И ЗначениеЗаполнено(ВыборкаМероприятий.ПредседательКомиссии) Тогда 
				ОбластьПодписьПредседателя = Макет.ПолучитьОбласть("ПодписьПредседателя");
				ОбластьПодписьПредседателя.Параметры.ПредседательКомиссии = ВыборкаМероприятий.ПредседательКомиссии;
				ТабДок.Вывести(ОбластьПодписьПредседателя);
			КонецЕсли;
			
			Если ИспользоватьКомиссии И ТЗЧленыКомиссии.Количество() Тогда
				ПерваяСтрока = Истина;
				ОбластьПодписьЧКСЗаголовком = Макет.ПолучитьОбласть("ПодписьЧКСЗаголовком");
				
				Для Каждого Строка Из ТЗЧленыКомиссии Цикл
					Если ПерваяСтрока Тогда
						ПерваяСтрока = Ложь;
						ОбластьПодписьЧКСЗаголовком.Параметры.ЧленКомиссии = Строка.ЧленКомиссии;
						ТабДок.Вывести(ОбластьПодписьЧКСЗаголовком);
						Продолжить;
					КонецЕсли;
					
					ОбластьПодписьЧК = Макет.ПолучитьОбласть("ПодписьЧК");
					ОбластьПодписьЧК.Параметры.ЧленКомиссии = Строка.ЧленКомиссии;
					ТабДок.Вывести(ОбластьПодписьЧК);
				КонецЦикла;
				
				Если ПерваяСтрока Тогда
					ТабДок.Вывести(ОбластьПодписьЧКСЗаголовком);
				КонецЕсли;
			КонецЕсли;
			
		КонецЦикла;
		
	КонецЦикла;
	
	ТабДок.ТолькоПросмотр = Истина;
	ТабДок.КлючПараметровПечати = "торо_ПечатьАктОВыполненииРегламентногоМероприятия";
	Возврат ТабДок;
	
КонецФункции

// Процедура выводит мероприятия акта в табличный документ.
Процедура ВывестиМероприятия(МассивСтрок,тзСписокРабот,СтрокаТаблицыРабот,ТабДок,МассивID,Ном,Комментарии,КолВо)
	Для каждого СписокРабот Из МассивСтрок Цикл
		СтрокаТаблицыРабот.Параметры.Номер = Ном;
		СтрокаТаблицыРабот.Параметры.Работа = СписокРабот.Работа;
		Если КолВо = Истина Тогда
    		СтрокаТаблицыРабот.Параметры.Количество=СписокРабот.Количество; 
		КонецЕсли;	
		Если Комментарии = Истина Тогда
			СтрокаТаблицыРабот.Параметры.Комментарий = СписокРабот.Комментарий;
		КонецЕсли;
		СтрокаТаблицыРабот.Параметры.ЕдИзм = СписокРабот.ЕдиницаИзмерения;
		ТабДок.Вывести(СтрокаТаблицыРабот);
		МассивID.Добавить(СписокРабот.ID);
		ВывестиМероприятия(тзСписокРабот.НайтиСтроки(Новый Структура("Родитель_ID",СписокРабот.ID)),тзСписокРабот,СтрокаТаблицыРабот,ТабДок,МассивID,Ном + ".1",Комментарии,КолВо);
		
		КоличествоРазделителей = СтрЧислоВхождений(Ном, ".");
		ПозицияКрайнегоРазделителя = СтрНайти(Ном, ".", , , ?(КоличествоРазделителей = 0, 1, КоличествоРазделителей));
		ЦелаяЧасть = Лев(Ном, ?(ПозицияКрайнегоРазделителя = 1, СтрДлина(Ном), ПозицияКрайнегоРазделителя));
		ДробнаяЧасть = Строка(Число(Сред(Ном, ПозицияКрайнегоРазделителя + 1)) + 1);
		Ном = ЦелаяЧасть + ДробнаяЧасть;
	КонецЦикла;
КонецПроцедуры

// СтандартныеПодсистемы.УправлениеДоступом

// См. УправлениеДоступомПереопределяемый.ПриЗаполненииСписковСОграничениемДоступа.
Процедура ПриЗаполненииОграниченияДоступа(Ограничение) Экспорт
	
	Ограничение.Текст =
	"РазрешитьЧтениеИзменение
	|ГДЕ
	|	ЗначениеРазрешено(Организация)
	|	И ЗначениеРазрешено(Подразделение)";

	Ограничение.ТекстДляВнешнихПользователей = Ограничение.Текст;
	
КонецПроцедуры

// Конец СтандартныеПодсистемы.УправлениеДоступом

#КонецОбласти

#КонецЕсли
