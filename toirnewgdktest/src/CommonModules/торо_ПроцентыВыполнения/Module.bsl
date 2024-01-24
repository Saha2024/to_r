/////////////////////////////////////////////////////////////////////////////////////
//
// торо_ПроцентыВыполнения: методы для работы с процентами выполнения ремонтных работ
//
/////////////////////////////////////////////////////////////////////////////////////

#Область СлужебныйПрограммныйИнтерфейс

Функция НайтиОперациюВРегистре(Операция, Родитель_ID, РемонтОборудования, НоваяОперация, Ссылка) Экспорт
	
	МассивВозврата = Новый Массив;
	СтруктураВозврата = Новый Структура("ID, Процент, Новая, Операция");
	
	Если НоваяОперация = Истина Тогда
		СтруктураВозврата.ID = Новый УникальныйИдентификатор;
		СтруктураВозврата.Процент = 0;
		СтруктураВозврата.Операция = Справочники.торо_ТехнологическиеОперации.ПустаяСсылка();
		СтруктураВозврата.Новая = Истина;
		МассивВозврата.Добавить(СтруктураВозврата);
	Иначе
		
		РезультатЗапроса = ПолучитьПроцентВыполненияОперации(Операция, Родитель_ID, РемонтОборудования, Ссылка);
		Если Не РезультатЗапроса.Пустой() Тогда
			Выборка = РезультатЗапроса.Выбрать();
			Пока Выборка.Следующий() Цикл
				СтруктураВозврата = Новый Структура("ID, Процент, Операция, Документ, НомерСтроки, Новая");
				ЗаполнитьЗначенияСвойств(СтруктураВозврата, Выборка);
				СтруктураВозврата.Новая = Ложь;
				МассивВозврата.Добавить(СтруктураВозврата);
			КонецЦикла; 
		Иначе
			СтруктураВозврата.ID = Новый УникальныйИдентификатор;
			СтруктураВозврата.Процент = 0;
			СтруктураВозврата.Новая = Истина;
			МассивВозврата.Добавить(СтруктураВозврата);
		КонецЕсли;
	КонецЕсли;
	Возврат МассивВозврата;
КонецФункции

Процедура РаспроведениеПлановыхЗатратБезАктов(СтарыйОбъект, НовыйОбъект, Регламентное) Экспорт
	
	Запрос = Новый Запрос;
	Запрос.МенеджерВременныхТаблиц = Новый МенеджерВременныхТаблиц;
	Запрос.Текст = 
		"ВЫБРАТЬ
		|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ID КАК ID
		|ПОМЕСТИТЬ ВТ_СтарыеРемонты
		|ИЗ
		|	Документ.торо_АктОВыполненииЭтапаРабот.РемонтыОборудования КАК торо_АктОВыполненииЭтапаРаботРемонтыОборудования
		|ГДЕ
		|	торо_АктОВыполненииЭтапаРаботРемонтыОборудования.Ссылка = &Ссылка
		|
		|ОБЪЕДИНИТЬ ВСЕ
		|
		|ВЫБРАТЬ
		|	торо_АктОВыполненииРегламентногоМероприятияМероприятия.ID
		|ИЗ
		|	Документ.торо_АктОВыполненииРегламентногоМероприятия.Мероприятия КАК торо_АктОВыполненииРегламентногоМероприятияМероприятия
		|ГДЕ
		|	торо_АктОВыполненииРегламентногоМероприятияМероприятия.Ссылка = &Ссылка
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	Ремонты.ID КАК ID
		|ПОМЕСТИТЬ ВТ_НовыеРемонты
		|ИЗ
		|	&НовыйОбъект КАК Ремонты
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	ВТ_СтарыеРемонты.ID КАК ID
		|ПОМЕСТИТЬ ВТ_УдаленныеРемонты
		|ИЗ
		|	ВТ_СтарыеРемонты КАК ВТ_СтарыеРемонты
		|		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_НовыеРемонты КАК ВТ_НовыеРемонты
		|		ПО ВТ_СтарыеРемонты.ID = ВТ_НовыеРемонты.ID
		|ГДЕ
		|	ВТ_НовыеРемонты.ID ЕСТЬ NULL
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	ВТ_УдаленныеРемонты.ID КАК ID,
		|	торо_ПлановыеЗатраты.Ссылка КАК Ссылка
		|ПОМЕСТИТЬ ВТ_ПлановыеЗатратыНаУдаление
		|ИЗ
		|	ВТ_УдаленныеРемонты КАК ВТ_УдаленныеРемонты
		|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.торо_ПлановыеЗатраты КАК торо_ПлановыеЗатраты
		|		ПО ВТ_УдаленныеРемонты.ID = торо_ПлановыеЗатраты.РемонтыОборудования_ID
		|ГДЕ
		|	торо_ПлановыеЗатраты.Проведен = ИСТИНА
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	ВТ_ПлановыеЗатратыНаУдаление.Ссылка КАК Ссылка,
		|	КОЛИЧЕСТВО(РАЗЛИЧНЫЕ ВЫБОР
		|			КОГДА торо_АктОВыполненииРегламентногоМероприятияМероприятия.Ссылка ЕСТЬ NULL
		|				ТОГДА торо_АктОВыполненииЭтапаРаботРемонтыОборудования.Ссылка
		|			ИНАЧЕ торо_АктОВыполненииРегламентногоМероприятияМероприятия.Ссылка
		|		КОНЕЦ) КАК КоличествоАктов
		|ПОМЕСТИТЬ ВТ_ПлановыеЗатратыСАктами
		|ИЗ
		|	ВТ_ПлановыеЗатратыНаУдаление КАК ВТ_ПлановыеЗатратыНаУдаление
		|		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_АктОВыполненииЭтапаРабот.РемонтыОборудования КАК торо_АктОВыполненииЭтапаРаботРемонтыОборудования
		|		ПО ВТ_ПлановыеЗатратыНаУдаление.ID = торо_АктОВыполненииЭтапаРаботРемонтыОборудования.ID
		|		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_АктОВыполненииРегламентногоМероприятия.Мероприятия КАК торо_АктОВыполненииРегламентногоМероприятияМероприятия
		|		ПО ВТ_ПлановыеЗатратыНаУдаление.ID = торо_АктОВыполненииРегламентногоМероприятияМероприятия.ID
		|
		|СГРУППИРОВАТЬ ПО
		|	ВТ_ПлановыеЗатратыНаУдаление.Ссылка
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	ВТ_ПлановыеЗатратыСАктами.Ссылка КАК Ссылка
		|ИЗ
		|	ВТ_ПлановыеЗатратыСАктами КАК ВТ_ПлановыеЗатратыСАктами
		|ГДЕ
		|	ВТ_ПлановыеЗатратыСАктами.КоличествоАктов = 1";
	
	Запрос.УстановитьПараметр("Ссылка", СтарыйОбъект);
	Если Регламентное Тогда
	 	Запрос.УстановитьПараметр("НовыйОбъект", НовыйОбъект.Мероприятия.Выгрузить(, "ID"));
	Иначе
		Запрос.УстановитьПараметр("НовыйОбъект", НовыйОбъект.РемонтыОборудования.Выгрузить(, "ID"));
	КонецЕсли;
	
	РезультатЗапроса = Запрос.Выполнить();
	
	ВыборкаДетальныеЗаписи = РезультатЗапроса.Выбрать();
	
	Пока ВыборкаДетальныеЗаписи.Следующий() Цикл
		ДокументЗатрат = ВыборкаДетальныеЗаписи.Ссылка.ПолучитьОбъект();
		ДокументЗатрат.Записать(РежимЗаписиДокумента.ОтменаПроведения);
	КонецЦикла;
	
КонецПроцедуры

Процедура СоздатьДокументыПлановыхЗатрат(РемонтыОборудования, РемонтныеРаботы, Регламентное) Экспорт

	Для каждого Строка Из РемонтыОборудования Цикл
		
		Если (Регламентное ИЛИ ОтсутствиеСметыПоРемонтуИзППР(Строка.ID)) Тогда
			
			// Получить список ТК для ремонта.
			// Для ТК, которые есть и в списке, и в ремонтных работах записываем данные в документ, иначе тк в плановые затраты не входит.
			ТаблицаТК = ПолучитьТаблицуТехКартРемонта(Строка.ОбъектРемонта, Строка.ВидРемонтныхРабот, Строка.ДатаНачала);

			НайденныйДокумент = Документы.торо_ПлановыеЗатраты.НайтиПоРеквизиту("РемонтыОборудования_ID", Строка.ID);
			
			Если ЗначениеЗаполнено(НайденныйДокумент) Тогда
				
				ДокументОбъект = НайденныйДокумент.ПолучитьОбъект();
				ТКОчищены = Ложь;
				
				Для каждого Работа Из РемонтныеРаботы Цикл  
					
					Если ТипЗнч(Работа.РемонтнаяРабота) = Тип("СправочникСсылка.торо_ИдентификаторыТехКарт") и Работа.РемонтыОборудования_ID = Строка.ID Тогда
						
						ПараметрыОтбора = Новый Структура("ТехКарта", Работа.РемонтнаяРабота); 
						НайденныеСтроки = ДокументОбъект.ТехКарты.НайтиСтроки(ПараметрыОтбора);
						НайденныеСтрокиИзНормативов = ТаблицаТК.НайтиСтроки(ПараметрыОтбора);
						Если НайденныеСтроки.Количество() = 0 И НайденныеСтрокиИзНормативов.Количество() <> 0 Тогда
							ДокументОбъект.ТехКарты.Очистить();
							ТКОчищены = Истина;
							Прервать;
						КонецЕсли; 
					КонецЕсли; 
				КонецЦикла;
				
				Если ТКОчищены Тогда
					Для каждого Работа Из РемонтныеРаботы Цикл
						Если ТипЗнч(Работа.РемонтнаяРабота) = Тип("СправочникСсылка.торо_ИдентификаторыТехКарт") и Работа.РемонтыОборудования_ID = Строка.ID Тогда
							
							ПараметрыОтбора = Новый Структура("ТехКарта", Работа.РемонтнаяРабота); 
							НайденныеСтрокиИзНормативов = ТаблицаТК.НайтиСтроки(ПараметрыОтбора);
							НайденныеСтрокиВДокументе = ДокументОбъект.ТехКарты.НайтиСтроки(ПараметрыОтбора);
							Если НайденныеСтрокиИзНормативов.Количество() <> 0 И НайденныеСтрокиИзНормативов[0].Количество > НайденныеСтрокиВДокументе.Количество() Тогда
								НоваяСтрока = ДокументОбъект.ТехКарты.Добавить();
								НоваяСтрока.ТехКарта = Работа.РемонтнаяРабота;
								НоваяСтрока.ТехКарта_ID = Работа.ID;
							КонецЕсли; 
						КонецЕсли;
					КонецЦикла;				
				КонецЕсли;  
				
				ДокументОбъект.Записать(РежимЗаписиДокумента.Проведение);
			Иначе   
								
				// Определим документ плана затрат. Если у источника есть и заполнено основание возьмем его,
				// если нет подставим сам документ источник.
				СтрокаРемонта = Неопределено;
				Если ТипЗнч(Строка.ДокументИсточник) = Тип("ДокументСсылка.торо_НарядНаРегламентноеМероприятие") 
					ИЛИ ТипЗнч(Строка.ДокументИсточник) = Тип("ДокументСсылка.торо_НарядНаВыполнениеРемонтныхРабот") Тогда  
					
					Ремонты = ?(Регламентное, Строка.ДокументИсточник.РегламентныеМероприятия, Строка.ДокументИсточник.РемонтыОборудования);
		                 
					СтрокаРемонта = Ремонты.Найти(Строка.ID, "ID");
				КонецЕсли;
				
				ДокументИсточник = ?(СтрокаРемонта = Неопределено, Строка.ДокументИсточник, СтрокаРемонта.ДокументИсточник); 
				
				ЕстьРекизитДокументОснование = 
					ОбщегоНазначения.ЕстьРеквизитОбъекта("ДокументОснование", ДокументИсточник.Метаданные());

				Если ЕстьРекизитДокументОснование Тогда
					ОснованиеИсточника = ОбщегоНазначения.ЗначениеРеквизитаОбъекта(ДокументИсточник, "ДокументОснование");
					ДокументОснование = ?(ЗначениеЗаполнено(ОснованиеИсточника), ОснованиеИсточника, ДокументИсточник);
				Иначе
					ДокументОснование = ДокументИсточник;
				КонецЕсли;		
				
				НовыйДокумент = Документы.торо_ПлановыеЗатраты.СоздатьДокумент();
				НовыйДокумент.Дата = Строка.ДатаНачала; 
				НовыйДокумент.ДокументПлана = ДокументОснование;
				НовыйДокумент.ОбъектРемонта = Строка.ОбъектРемонта;
				НовыйДокумент.РемонтыОборудования_ID = Строка.ID;    
				
				Для каждого Работа Из РемонтныеРаботы Цикл  
					
					Если ТипЗнч(Работа.РемонтнаяРабота) = Тип("СправочникСсылка.торо_ИдентификаторыТехКарт") и Работа.РемонтыОборудования_ID = Строка.ID Тогда 
						
						ПараметрыОтбора = Новый Структура("ТехКарта", Работа.РемонтнаяРабота); 
						НайденныеСтрокиИзНормативов = ТаблицаТК.НайтиСтроки(ПараметрыОтбора);
						НайденныеСтрокиВДокументе = НовыйДокумент.ТехКарты.НайтиСтроки(ПараметрыОтбора);
						Если НайденныеСтрокиИзНормативов.Количество() <> 0 И НайденныеСтрокиИзНормативов[0].Количество > НайденныеСтрокиВДокументе.Количество() Тогда
							НоваяСтрока = НовыйДокумент.ТехКарты.Добавить();
							НоваяСтрока.ТехКарта = Работа.РемонтнаяРабота;
							НоваяСтрока.ТехКарта_ID = Работа.ID;
						КонецЕсли;
					КонецЕсли; 
				КонецЦикла; 
				
				НовыйДокумент.Записать(РежимЗаписиДокумента.Проведение);	
			КонецЕсли; 
		КонецЕсли; 
	КонецЦикла; 
	
КонецПроцедуры

Функция ОтсутствиеСметыПоРемонтуИзППР(IDРемонта)

	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
	               |	торо_ПлановыеРемонтныеРаботыСрезПоследних.ID КАК ID
	               |ПОМЕСТИТЬ ВТ_ППР
	               |ИЗ
	               |	РегистрСведений.торо_ПлановыеРемонтныеРаботы.СрезПоследних(, ID = &IDРемонта) КАК торо_ПлановыеРемонтныеРаботыСрезПоследних
	               |ГДЕ
	               |	НЕ торо_ПлановыеРемонтныеРаботыСрезПоследних.Отменен
	               |
	               |ИНДЕКСИРОВАТЬ ПО
	               |	ID
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ РАЗЛИЧНЫЕ
	               |	торо_ЗаявкиПоРемонтамСрезПоследних.IDРемонта КАК IDРемонта
	               |ПОМЕСТИТЬ ВТ_Заявки
	               |ИЗ
	               |	РегистрСведений.торо_ЗаявкиПоРемонтам.СрезПоследних(, IDРемонта = &IDРемонта) КАК торо_ЗаявкиПоРемонтамСрезПоследних
	               |ГДЕ
	               |	НЕ торо_ЗаявкиПоРемонтамСрезПоследних.Отменен
	               |
	               |ИНДЕКСИРОВАТЬ ПО
	               |	IDРемонта
	               |;
	               |
				   //|////////////////////////////////////////////////////////////////////////////////
				   //|ВЫБРАТЬ
				   //|	торо_ОтмененныеРемонтыСрезПоследних.ID КАК ID
				   //|ПОМЕСТИТЬ ВТ_Отмененные
				   //|ИЗ
				   //|	РегистрСведений.торо_ОтмененныеРемонты.СрезПоследних(, ID = &IDРемонта) КАК торо_ОтмененныеРемонтыСрезПоследних
				   //|ГДЕ
				   //|	торо_ОтмененныеРемонтыСрезПоследних.Отменен = ИСТИНА
				   //|
				   //|ИНДЕКСИРОВАТЬ ПО
				   //|	ID
				   //|;
				   //|
				   //|////////////////////////////////////////////////////////////////////////////////
				   //|ВЫБРАТЬ
				   //|	ВТ_Заявки.IDРемонта КАК IDРемонта
				   //|ПОМЕСТИТЬ ВТ_НеОтмененныеЗаявки
				   //|ИЗ
				   //|	ВТ_Заявки КАК ВТ_Заявки
				   //|		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_Отмененные КАК ВТ_Отмененные
				   //|		ПО ВТ_Заявки.IDРемонта = ВТ_Отмененные.ID
				   //|ГДЕ
				   //|	ВТ_Отмененные.ID ЕСТЬ NULL
				   //|
				   //|ИНДЕКСИРОВАТЬ ПО
				   //|	IDРемонта
				   //|;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_ППР.ID КАК ID
	               |ИЗ
	               |	ВТ_ППР КАК ВТ_ППР
	               |		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_Заявки КАК ВТ_Заявки
	               |		ПО (ВТ_ППР.ID = ВТ_Заявки.IDРемонта)
	               |ГДЕ
	               |	ВТ_Заявки.IDРемонта ЕСТЬ NULL";
	
	Запрос.УстановитьПараметр("IDРемонта", IDРемонта);
	
	РезультатЗапроса = Запрос.Выполнить();
	
	Возврат Не РезультатЗапроса.Пустой();
	
КонецФункции 

Процедура ПолучитьНевыполненныеПроценты(ТЧ_РемонтныхРабот, Ссылка, РемонтыОборудования_ID) Экспорт
	
	Запрос = Новый Запрос;
	Запрос.МенеджерВременныхТаблиц = Новый МенеджерВременныхТаблиц();
	
	Запрос.Текст = "ВЫБРАТЬ
	               |	ТаблицаРемонтныхРабот.ID КАК ID,
	               |	ТаблицаРемонтныхРабот.РемонтнаяРабота КАК РемонтнаяРабота,
	               |	ТаблицаРемонтныхРабот.Родитель_ID КАК Родитель_ID,
	               |	ТаблицаРемонтныхРабот.Количество КАК Количество,
	               |	ТаблицаРемонтныхРабот.РемонтыОборудования_ID КАК РемонтыОборудования_ID
	               |ПОМЕСТИТЬ ВТ_РемонтыеРаботы
	               |ИЗ
	               |	&ТаблицаРемонтныхРабот КАК ТаблицаРемонтныхРабот
	               |
	               |ИНДЕКСИРОВАТЬ ПО
	               |	РемонтыОборудования_ID
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	торо_ПроцентВыполненныхРабот.IDОперации КАК ID,
	               |	торо_ПроцентВыполненныхРабот.Операция КАК Операция,
	               |	торо_ПроцентВыполненныхРабот.Регистратор КАК Документ,
	               |	торо_ПроцентВыполненныхРабот.Процент КАК Процент,
				   |	торо_ПроцентВыполненныхРабот.НомерСтроки КАК НомерСтроки,
	               |	ВЫБОР
	               |		КОГДА ТИПЗНАЧЕНИЯ(торо_ПроцентВыполненныхРабот.Регистратор) = ТИП(Документ.торо_ПлановыеЗатраты)
	               |			ТОГДА 0
	               |		КОГДА ТИПЗНАЧЕНИЯ(торо_ПроцентВыполненныхРабот.Регистратор) = ТИП(Документ.торо_ЗаявкаНаРемонт)
	               |			ТОГДА 1
	               |		КОГДА ТИПЗНАЧЕНИЯ(торо_ПроцентВыполненныхРабот.Регистратор) = ТИП(Документ.торо_ЗакрытиеЗаявокИРемонтов)
	               |			ТОГДА 2
	               |		КОГДА ТИПЗНАЧЕНИЯ(торо_ПроцентВыполненныхРабот.Регистратор) = ТИП(Документ.торо_НарядНаВыполнениеРемонтныхРабот)
	               |			ТОГДА 3
	               |		КОГДА ТИПЗНАЧЕНИЯ(торо_ПроцентВыполненныхРабот.Регистратор) = ТИП(Документ.торо_НарядНаРегламентноеМероприятие)
	               |			ТОГДА 4
	               |		ИНАЧЕ 5
	               |	КОНЕЦ КАК Приоритет,
	               |	торо_ПроцентВыполненныхРабот.IDРодителя КАК IDРодителя,
	               |	торо_ПроцентВыполненныхРабот.IDРемонта КАК IDРемонта
	               |ПОМЕСТИТЬ ВТ_Операции
	               |ИЗ
	               |	РегистрНакопления.торо_ПроцентВыполненныхРабот КАК торо_ПроцентВыполненныхРабот
	               |ГДЕ
	               |	торо_ПроцентВыполненныхРабот.IDРемонта В
	               |			(ВЫБРАТЬ
	               |				ВТ_РемонтыеРаботы.РемонтыОборудования_ID КАК РемонтыОборудования_ID
	               |			ИЗ
	               |				ВТ_РемонтыеРаботы КАК ВТ_РемонтыеРаботы)
	               |	И торо_ПроцентВыполненныхРабот.Регистратор <> &Ссылка
	               |	И торо_ПроцентВыполненныхРабот.IDРемонта = &РемонтыОборудования_ID
				   |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
				   |   ВТ_Операции.ID КАК ID,
				   |   ВТ_Операции.Операция КАК Операция,
				   |   ВТ_Операции.Документ КАК Документ,
				   |   ВТ_Операции.Процент КАК Процент,
				   |   ВТ_Операции.Приоритет КАК Приоритет, 
				   |   ВТ_Операции.IDРодителя КАК IDРодителя,
				   |   ВТ_Операции.IDРемонта КАК IDРемонта
				   |ПОМЕСТИТЬ ВТ_ВсеОперации
				   |ИЗ
				   |  ВТ_Операции КАК ВТ_Операции
				   |
				   |ОБЪЕДИНИТЬ 
				   | 
				   |ВЫБРАТЬ 
				   |  торо_ОперацииВнешнихОснованийДляРабот.ID,
				   |  торо_ОперацииВнешнихОснованийДляРабот.РемонтнаяРабота,
				   |  торо_ОперацииВнешнихОснованийДляРабот.Регистратор, 
				   |  торо_ОперацииВнешнихОснованийДляРабот.Количество*100,
				   |  0, 
				   |  торо_ОперацииВнешнихОснованийДляРабот.Родитель_ID, 
				   |  торо_ОперацииВнешнихОснованийДляРабот.Предписание_ID
				   |ИЗ 
				   |  РегистрСведений.торо_ОперацииВнешнихОснованийДляРабот КАК торо_ОперацииВнешнихОснованийДляРабот
				   |        ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_Операции КАК ВТ_Операции
				   |        ПО торо_ОперацииВнешнихОснованийДляРабот.ID = ВТ_Операции.ID
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_ВсеОперации.ID КАК ID,
	               |	МИНИМУМ(ВТ_ВсеОперации.Приоритет) КАК Приоритет,
	               |	ВТ_ВсеОперации.IDРемонта КАК IDРемонта
	               |ПОМЕСТИТЬ ВТ_МинимумПриоритетов
	               |ИЗ
	               |	ВТ_ВсеОперации КАК ВТ_ВсеОперации
	               |ГДЕ
	               |	ВТ_ВсеОперации.Приоритет < 5
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТ_ВсеОперации.ID,
	               |	ВТ_ВсеОперации.IDРемонта
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_ВсеОперации.ID КАК ID,
	               |	МАКСИМУМ(ВТ_ВсеОперации.Процент) КАК Процент,
	               |	ВТ_ВсеОперации.Приоритет КАК Приоритет,
	               |	ВТ_ВсеОперации.IDРемонта КАК IDРемонта
	               |ПОМЕСТИТЬ ВТ_План
	               |ИЗ
	               |	ВТ_ВсеОперации КАК ВТ_ВсеОперации
	               |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_МинимумПриоритетов КАК ВТ_МинимумПриоритетов
	               |		ПО ВТ_ВсеОперации.ID = ВТ_МинимумПриоритетов.ID
	               |			И ВТ_ВсеОперации.Приоритет = ВТ_МинимумПриоритетов.Приоритет
	               |			И ВТ_ВсеОперации.IDРемонта = ВТ_МинимумПриоритетов.IDРемонта
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТ_ВсеОперации.ID,
	               |	ВТ_ВсеОперации.Приоритет,
	               |	ВТ_ВсеОперации.IDРемонта
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_ВсеОперации.ID КАК ID,
	               |	СУММА(ВТ_ВсеОперации.Процент) КАК Процент,
	               |	ВТ_ВсеОперации.IDРемонта КАК IDРемонта
	               |ПОМЕСТИТЬ ВТ_Факт
	               |ИЗ
	               |	ВТ_ВсеОперации КАК ВТ_ВсеОперации
	               |ГДЕ
	               |	ВТ_ВсеОперации.Приоритет = 5
	               |	И ВТ_ВсеОперации.Процент < 0
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТ_ВсеОперации.ID,
	               |	ВТ_ВсеОперации.IDРемонта
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_ВсеОперации.ID КАК ID,
	               |	ВТ_ВсеОперации.Операция КАК Операция,
	               |	ЕСТЬNULL(ВТ_План.Процент, 0) + ЕСТЬNULL(ВТ_Факт.Процент, 0) КАК Процент,
	               |	ВТ_ВсеОперации.IDРодителя КАК IDРодителя,
	               |	ВТ_ВсеОперации.IDРемонта КАК IDРемонта
	               |ПОМЕСТИТЬ ВТ_ПлановоеОсталось
	               |ИЗ
	               |	ВТ_ВсеОперации КАК ВТ_ВсеОперации
	               |		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_План КАК ВТ_План
	               |		ПО ВТ_ВсеОперации.ID = ВТ_План.ID
	               |			И ВТ_ВсеОперации.Приоритет = ВТ_План.Приоритет
	               |			И ВТ_ВсеОперации.IDРемонта = ВТ_План.IDРемонта
	               |		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_Факт КАК ВТ_Факт
	               |		ПО ВТ_ВсеОперации.ID = ВТ_Факт.ID
	               |			И ВТ_ВсеОперации.IDРемонта = ВТ_Факт.IDРемонта
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ РАЗЛИЧНЫЕ
	               |	ВТ_РемонтыеРаботы.ID КАК ID,
	               |	ВТ_РемонтыеРаботы.РемонтнаяРабота КАК РемонтнаяРабота,
	               |	ВТ_РемонтыеРаботы.Родитель_ID КАК Родитель_ID,
	               |	ВТ_РемонтыеРаботы.РемонтыОборудования_ID КАК РемонтыОборудования_ID,
				  // |    (ВТ_ПлановоеОсталось.Процент ЕСТЬ NULL) КАК НоваяОперация,
	               |	ВЫБОР
	               |		КОГДА ВТ_ПлановоеОсталось.Процент ЕСТЬ NULL
	               |			ТОГДА 100
	               |		КОГДА ВТ_ПлановоеОсталось.Процент < 0
	               |			ТОГДА (100 + ВТ_ПлановоеОсталось.Процент) / ВТ_РемонтыеРаботы.Количество
	               |		ИНАЧЕ ВТ_ПлановоеОсталось.Процент / ВТ_РемонтыеРаботы.Количество
	               |	КОНЕЦ КАК Процент
	               |ИЗ
	               |	ВТ_РемонтыеРаботы КАК ВТ_РемонтыеРаботы
	               |		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_ПлановоеОсталось КАК ВТ_ПлановоеОсталось
	               |		ПО ВТ_РемонтыеРаботы.ID = ВТ_ПлановоеОсталось.ID
	               |			И ВТ_РемонтыеРаботы.РемонтыОборудования_ID = ВТ_ПлановоеОсталось.IDРемонта
	               |			И ВТ_РемонтыеРаботы.Родитель_ID = ВТ_ПлановоеОсталось.IDРодителя
				   | ГДЕ ВТ_РемонтыеРаботы.РемонтыОборудования_ID = &РемонтыОборудования_ID";
	Если ТипЗнч(ТЧ_РемонтныхРабот) = Тип("ТаблицаЗначений") Тогда
	    Запрос.УстановитьПараметр("ТаблицаРемонтныхРабот", ТЧ_РемонтныхРабот);
	Иначе
		Запрос.УстановитьПараметр("ТаблицаРемонтныхРабот", ТЧ_РемонтныхРабот.Выгрузить());
	КонецЕсли; 
	
	Запрос.УстановитьПараметр("Ссылка", Ссылка);
	Если НЕ ЗначениеЗаполнено(РемонтыОборудования_ID) Тогда
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "И торо_ПроцентВыполненныхРабот.IDРемонта = &РемонтыОборудования_ID", "");
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "ВТ_РемонтыеРаботы.РемонтыОборудования_ID = &РемонтыОборудования_ID", "Истина");
	Иначе
		Запрос.УстановитьПараметр("РемонтыОборудования_ID", РемонтыОборудования_ID);
	КонецЕсли; 
	РезультатЗапроса = Запрос.Выполнить();
	Если Не РезультатЗапроса.Пустой() Тогда
		СтруктураПоиска = Новый Структура("ID, РемонтнаяРабота, Родитель_ID, РемонтыОборудования_ID");
		Выборка = РезультатЗапроса.Выбрать();
		Пока Выборка.Следующий() Цикл
			ЗаполнитьЗначенияСвойств(СтруктураПоиска, Выборка);
			Найденное = ТЧ_РемонтныхРабот.НайтиСтроки(СтруктураПоиска);
			Если Найденное.Количество() = 1 Тогда
				Найденное[0].ПроцентВыполненияРабот = Выборка.Процент;
			КонецЕсли; 
		КонецЦикла; 
	КонецЕсли;
КонецПроцедуры

Функция ПодобратьIDТКИзПлановыхЗатрат(РемонтыОборудования_ID, ТехКарта) Экспорт
	
	ТехКарта_ID = Новый УникальныйИдентификатор;
	Запрос = Новый Запрос;
	Запрос.Текст = 
	"ВЫБРАТЬ
	|	торо_ПлановыеЗатратыТехКарты.ТехКарта_ID КАК ТехКарта_ID
	|ИЗ
	|	Документ.торо_ПлановыеЗатраты.ТехКарты КАК торо_ПлановыеЗатратыТехКарты
	|ГДЕ
	|	торо_ПлановыеЗатратыТехКарты.ТехКарта = &ТехКарта
	|	И торо_ПлановыеЗатратыТехКарты.Ссылка.РемонтыОборудования_ID = &РемонтыОборудования_ID";
	
	Запрос.УстановитьПараметр("РемонтыОборудования_ID", РемонтыОборудования_ID);
	Запрос.УстановитьПараметр("ТехКарта", ТехКарта);
	
	РезультатЗапроса = Запрос.Выполнить();
	
	Выборка = РезультатЗапроса.Выбрать();
	
	Если Выборка.Следующий() Тогда
		ТехКарта_ID = Выборка.ТехКарта_ID;	
	КонецЕсли;
	
	Возврат ТехКарта_ID;
	
КонецФункции

Функция ОставшийсяПроцентВыполнения(Операция, РемонтыОборудования_ID, Родитель_ID, IDРемонта) Экспорт	
	РезультатЗапроса = ПолучитьПроцентВыполненияОперации(Операция, Родитель_ID, РемонтыОборудования_ID, Документы.торо_АктОВыполненииРегламентногоМероприятия.ПустаяСсылка(), IDРемонта);
	Если Не РезультатЗапроса.Пустой() Тогда
		Выборка = РезультатЗапроса.Выбрать();
		Пока Выборка.Следующий() Цикл
			Возврат Выборка.Процент;
		КонецЦикла;
	Иначе	
		Возврат 100;	
	КонецЕсли;
КонецФункции

Функция ПолучитьПроцентВыполненияОперации(Операция, Родитель_ID, РемонтОборудования, Ссылка, IDРемонта = Неопределено) Экспорт 
	Запрос = Новый Запрос;
	Запрос.МенеджерВременныхТаблиц = Новый МенеджерВременныхТаблиц;
	Запрос.Текст = "ВЫБРАТЬ
	               |	торо_ПроцентВыполненныхРабот.IDОперации КАК ID,
	               |	торо_ПроцентВыполненныхРабот.Операция КАК Операция,
	               |	торо_ПроцентВыполненныхРабот.Регистратор КАК Документ,
	               |	торо_ПроцентВыполненныхРабот.Процент КАК Процент,
	               |	торо_ПроцентВыполненныхРабот.НомерСтроки КАК НомерСтроки,
	               |	ВЫБОР
	               |		КОГДА ТИПЗНАЧЕНИЯ(торо_ПроцентВыполненныхРабот.Регистратор) = ТИП(Документ.торо_ПлановыеЗатраты)
	               |			ТОГДА 0
	               |		КОГДА ТИПЗНАЧЕНИЯ(торо_ПроцентВыполненныхРабот.Регистратор) = ТИП(Документ.торо_ЗаявкаНаРемонт)
	               |			ТОГДА 1
	               |		КОГДА ТИПЗНАЧЕНИЯ(торо_ПроцентВыполненныхРабот.Регистратор) = ТИП(Документ.торо_ЗакрытиеЗаявокИРемонтов)
	               |			ТОГДА 2
	               |		КОГДА ТИПЗНАЧЕНИЯ(торо_ПроцентВыполненныхРабот.Регистратор) = ТИП(Документ.торо_НарядНаВыполнениеРемонтныхРабот)
	               |			ТОГДА 3
	               |		КОГДА ТИПЗНАЧЕНИЯ(торо_ПроцентВыполненныхРабот.Регистратор) = ТИП(Документ.торо_НарядНаРегламентноеМероприятие)
	               |			ТОГДА 4
	               |		ИНАЧЕ 5
	               |	КОНЕЦ КАК Приоритет,
	               |	торо_ПроцентВыполненныхРабот.IDРодителя КАК IDРодителя,
	               |	торо_ПроцентВыполненныхРабот.Регистратор.Дата КАК РегистраторДата
				   |ПОМЕСТИТЬ ВТ_Операции
	               |ИЗ
	               |	РегистрНакопления.торо_ПроцентВыполненныхРабот КАК торо_ПроцентВыполненныхРабот
	               |ГДЕ
	               |	торо_ПроцентВыполненныхРабот.IDРемонта = &РемонтыОборудования_ID
	               |	И торо_ПроцентВыполненныхРабот.IDРодителя = &Родитель_ID
	               |	И торо_ПроцентВыполненныхРабот.Операция = &РемонтнаяРабота
	               |	И торо_ПроцентВыполненныхРабот.Регистратор <> &Ссылка
				   |;
				   |
				   |////////////////////////////////////////////////////////////////////////////////
				   |ВЫБРАТЬ
				   |   ВТ_Операции.ID КАК ID,
				   |   ВТ_Операции.Операция КАК Операция,
				   |   ВТ_Операции.Документ КАК Документ,
				   |   ВТ_Операции.Процент КАК Процент,
				   |   ВТ_Операции.Приоритет КАК Приоритет, 
				   |   ВТ_Операции.НомерСтроки КАК НомерСтроки,
				   |   ВТ_Операции.IDРодителя КАК IDРодителя,
				   |   ВТ_Операции.РегистраторДата КАК РегистраторДата
				   |ПОМЕСТИТЬ ВТ_ВсеОперации
				   |ИЗ
				   |  ВТ_Операции КАК ВТ_Операции
				   |
				   |ОБЪЕДИНИТЬ 
				   | 
				   |ВЫБРАТЬ 
				   |  торо_ОперацииВнешнихОснованийДляРабот.ID,
				   |  торо_ОперацииВнешнихОснованийДляРабот.РемонтнаяРабота,
				   |  торо_ОперацииВнешнихОснованийДляРабот.Регистратор, 
				   |  торо_ОперацииВнешнихОснованийДляРабот.Количество*100,
				   |  0, 
				   |  торо_ОперацииВнешнихОснованийДляРабот.НомерСтроки,
				   |  торо_ОперацииВнешнихОснованийДляРабот.Родитель_ID, 
				   |  торо_ОперацииВнешнихОснованийДляРабот.Регистратор.Дата
				   |ИЗ 
				   |  РегистрСведений.торо_ОперацииВнешнихОснованийДляРабот КАК торо_ОперацииВнешнихОснованийДляРабот
				   |        ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_Операции КАК ВТ_Операции
				   |        ПО торо_ОперацииВнешнихОснованийДляРабот.ID = ВТ_Операции.ID
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_ВсеОперации.ID КАК ID,
	               |	МИНИМУМ(ВТ_ВсеОперации.Приоритет) КАК Приоритет
	               |ПОМЕСТИТЬ ВТ_МинимумПриоритетов
	               |ИЗ
	               |	ВТ_ВсеОперации КАК ВТ_ВсеОперации
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТ_ВсеОперации.ID
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_ВсеОперации.ID КАК ID,
	               |	ВТ_ВсеОперации.Операция КАК Операция,
	               |	ВТ_ВсеОперации.Документ КАК Документ,
	               |	ВТ_ВсеОперации.Процент КАК Процент,
				   |	ВТ_ВсеОперации.НомерСтроки КАК НомерСтроки,
	               |	ВТ_ВсеОперации.Приоритет КАК Приоритет,
	               |	ВТ_ВсеОперации.IDРодителя КАК IDРодителя,
	               |	ВТ_ВсеОперации.РегистраторДата КАК РегистраторДата
	               |ПОМЕСТИТЬ ВТ_ПервоеПоявлениеОперации
	               |ИЗ
	               |	ВТ_ВсеОперации КАК ВТ_ВсеОперации
	               |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_МинимумПриоритетов КАК ВТ_МинимумПриоритетов
	               |		ПО (ВТ_МинимумПриоритетов.ID = ВТ_ВсеОперации.ID)
	               |			И (ВТ_МинимумПриоритетов.Приоритет = ВТ_ВсеОперации.Приоритет)
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_ВсеОперации.ID КАК ID,
	               |	СУММА(ВТ_ВсеОперации.Процент) КАК Процент,
	               |	ВТ_ВсеОперации.Приоритет КАК Приоритет
	               |ПОМЕСТИТЬ ВТ_План
	               |ИЗ
	               |	ВТ_ВсеОперации КАК ВТ_ВсеОперации
	               |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_МинимумПриоритетов КАК ВТ_МинимумПриоритетов
	               |		ПО ВТ_ВсеОперации.ID = ВТ_МинимумПриоритетов.ID
	               |			И ВТ_ВсеОперации.Приоритет = ВТ_МинимумПриоритетов.Приоритет
	               |ГДЕ
	               |	ВТ_ВсеОперации.Приоритет <> 5
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТ_ВсеОперации.ID,
	               |	ВТ_ВсеОперации.Приоритет
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_ВсеОперации.ID КАК ID,
	               |	СУММА(ВТ_ВсеОперации.Процент) КАК Процент
	               |ПОМЕСТИТЬ ВТ_Факт
	               |ИЗ
	               |	ВТ_ВсеОперации КАК ВТ_ВсеОперации
	               |ГДЕ
	               |	ВТ_ВсеОперации.Приоритет = 5
	               |	И ВТ_ВсеОперации.Процент < 0
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТ_ВсеОперации.ID
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_ПервоеПоявлениеОперации.ID КАК ID,
	               |	ВТ_ПервоеПоявлениеОперации.Операция КАК Операция,
	               |	ВТ_ПервоеПоявлениеОперации.Документ КАК Документ,
				   |	ВТ_ПервоеПоявлениеОперации.НомерСтроки КАК НомерСтроки,
	               |	ВТ_ПервоеПоявлениеОперации.Приоритет КАК Приоритет,
	               |	ЕСТЬNULL(ВТ_План.Процент, 0) + ЕСТЬNULL(ВТ_Факт.Процент, 0) КАК Процент,
	               |	ВТ_ПервоеПоявлениеОперации.IDРодителя КАК IDРодителя,
	               |	ВТ_ПервоеПоявлениеОперации.РегистраторДата КАК РегистраторДата
	               |ПОМЕСТИТЬ ВТ_ВсеДокументы
	               |ИЗ
	               |	ВТ_ПервоеПоявлениеОперации КАК ВТ_ПервоеПоявлениеОперации
	               |		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_План КАК ВТ_План
	               |		ПО ВТ_ПервоеПоявлениеОперации.ID = ВТ_План.ID
	               |			И ВТ_ПервоеПоявлениеОперации.Приоритет = ВТ_План.Приоритет
	               |		ЛЕВОЕ СОЕДИНЕНИЕ ВТ_Факт КАК ВТ_Факт
	               |		ПО ВТ_ПервоеПоявлениеОперации.ID = ВТ_Факт.ID
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	МИНИМУМ(ВТ_ВсеДокументы.РегистраторДата) КАК РегистраторДата,
	               |	ВТ_ВсеДокументы.ID КАК ID,
	               |	ВТ_ВсеДокументы.Операция КАК Операция,
	               |	ВТ_ВсеДокументы.IDРодителя КАК IDРодителя
	               |ПОМЕСТИТЬ ВТ_МинимумПоДате
	               |ИЗ
	               |	ВТ_ВсеДокументы КАК ВТ_ВсеДокументы
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТ_ВсеДокументы.ID,
	               |	ВТ_ВсеДокументы.Операция,
	               |	ВТ_ВсеДокументы.IDРодителя
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_ВсеДокументы.ID КАК ID,
	               |	ВТ_ВсеДокументы.Операция КАК Операция,
	               |	ВТ_ВсеДокументы.Документ КАК Документ,
				   |	ВТ_ВсеДокументы.НомерСтроки КАК НомерСтроки,
	               |	ВТ_ВсеДокументы.Приоритет КАК Приоритет,
	               |	ВТ_ВсеДокументы.Процент КАК Процент,
	               |	ВТ_ВсеДокументы.IDРодителя КАК IDРодителя,
	               |	ВТ_ВсеДокументы.РегистраторДата КАК РегистраторДата
	               |ИЗ
	               |	ВТ_ВсеДокументы КАК ВТ_ВсеДокументы
	               |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_МинимумПоДате КАК ВТ_МинимумПоДате
	               |		ПО (ВТ_МинимумПоДате.РегистраторДата = ВТ_ВсеДокументы.РегистраторДата)
	               |			И (ВТ_МинимумПоДате.ID = ВТ_ВсеДокументы.ID)
	               |			И (ВТ_МинимумПоДате.IDРодителя = ВТ_ВсеДокументы.IDРодителя)";
	
	Запрос.УстановитьПараметр("РемонтнаяРабота", Операция);
	Запрос.УстановитьПараметр("РемонтыОборудования_ID", РемонтОборудования);
	Запрос.УстановитьПараметр("Родитель_ID", Родитель_ID);
	Запрос.УстановитьПараметр("Ссылка", Ссылка);	
	
	Если IDРемонта <> Неопределено Тогда
		Запрос.Текст = Запрос.Текст + " ГДЕ ВТ_ВсеДокументы.ID = &IDРемонта";
		Запрос.УстановитьПараметр("IDРемонта", Строка(IDРемонта));
	КонецЕсли;
	
	Возврат Запрос.Выполнить();

КонецФункции 

Функция ПолучитьТаблицуТехКартРемонта(ОбъектРемонта, ВидРемонта, ДатаНачала)
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
		"ВЫБРАТЬ
		|	торо_НормативныеРемонтыОборудования.ОбъектРемонта КАК ОбъектРемонта,
		|	торо_НормативныеРемонтыОборудования.ВидРемонта КАК ВидРемонта,
		|	торо_НормативныеРемонтыОборудования.НормативныйРемонт КАК НормативныйРемонт
		|ПОМЕСТИТЬ ВТ_НормативыОР
		|ИЗ
		|	РегистрСведений.торо_НормативныеРемонтыОборудования КАК торо_НормативныеРемонтыОборудования
		|ГДЕ
		|	торо_НормативныеРемонтыОборудования.ОбъектРемонта = &ОбъектРемонта
		|	И торо_НормативныеРемонтыОборудования.ВидРемонта = &ВидРемонта
		|	И торо_НормативныеРемонтыОборудования.Приоритет = ИСТИНА
		|
		|ОБЪЕДИНИТЬ ВСЕ
		|
		|ВЫБРАТЬ
		|	торо_РегламентныеМероприятияИСпискиОбъектов.СписокОбъектов,
		|	торо_РегламентныеМероприятияИСпискиОбъектов.ВидМероприятия,
		|	торо_РегламентныеМероприятияИСпискиОбъектов.НормативныйРемонт
		|ИЗ
		|	РегистрСведений.торо_РегламентныеМероприятияИСпискиОбъектов КАК торо_РегламентныеМероприятияИСпискиОбъектов
		|ГДЕ
		|	торо_РегламентныеМероприятияИСпискиОбъектов.СписокОбъектов = &ОбъектРемонта
		|	И торо_РегламентныеМероприятияИСпискиОбъектов.ВидМероприятия = &ВидРемонта
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	ВТ_НормативыОР.ОбъектРемонта КАК ОбъектРемонта,
		|	ВТ_НормативыОР.ВидРемонта КАК ВидРемонта,
		|	торо_ВерсииТехКартСрезПоследних.ИдентификаторТехКарты КАК ИдентификаторТехКарты,
		|	торо_ВерсииТехКартСрезПоследних.ТехКарта КАК ТехКарта
		|ПОМЕСТИТЬ ВТ_ВерсияТехКарты
		|ИЗ
		|	ВТ_НормативыОР КАК ВТ_НормативыОР
		|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.торо_ВерсииТехКарт.СрезПоследних(&ДатаНачала, ) КАК торо_ВерсииТехКартСрезПоследних
		|		ПО ВТ_НормативыОР.НормативныйРемонт = торо_ВерсииТехКартСрезПоследних.ИдентификаторТехКарты
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	ВТ_ВерсияТехКарты.ОбъектРемонта КАК ОбъектРемонта,
		|	ВТ_ВерсияТехКарты.ВидРемонта КАК ВидРемонта,
		|	торо_СоставИерархическихТехКартРемонтов.Состав КАК ТехКарта,
		|	торо_СоставИерархическихТехКартРемонтов.Количество КАК Количество,
		|	ВТ_ВерсияТехКарты.ИдентификаторТехКарты КАК ИдентификаторТехКарты
		|ПОМЕСТИТЬ ВТ_ВерсииСИдентификаторами
		|ИЗ
		|	ВТ_ВерсияТехКарты КАК ВТ_ВерсияТехКарты
		|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.торо_СоставИерархическихТехКартРемонтов КАК торо_СоставИерархическихТехКартРемонтов
		|		ПО ВТ_ВерсияТехКарты.ТехКарта = торо_СоставИерархическихТехКартРемонтов.ТехКарта
		|
		|ОБЪЕДИНИТЬ ВСЕ
		|
		|ВЫБРАТЬ
		|	ВТ_ВерсияТехКарты.ОбъектРемонта,
		|	ВТ_ВерсияТехКарты.ВидРемонта,
		|	ВТ_ВерсияТехКарты.ТехКарта,
		|	1,
		|	ВТ_ВерсияТехКарты.ИдентификаторТехКарты
		|ИЗ
		|	ВТ_ВерсияТехКарты КАК ВТ_ВерсияТехКарты
		|;
		|
		|////////////////////////////////////////////////////////////////////////////////
		|ВЫБРАТЬ
		|	ВТ_ВерсииСИдентификаторами.ОбъектРемонта КАК ОбъектРемонта,
		|	ВТ_ВерсииСИдентификаторами.ВидРемонта КАК ВидРемонта,
		|	ВТ_ВерсииСИдентификаторами.Количество КАК Количество,
		|	торо_ВерсииТехКарт.ИдентификаторТехКарты КАК ТехКарта
		|ИЗ
		|	ВТ_ВерсииСИдентификаторами КАК ВТ_ВерсииСИдентификаторами
		|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ РегистрСведений.торо_ВерсииТехКарт КАК торо_ВерсииТехКарт
		|		ПО ВТ_ВерсииСИдентификаторами.ТехКарта = торо_ВерсииТехКарт.ТехКарта";
	
	Запрос.УстановитьПараметр("ВидРемонта", ВидРемонта);
	Запрос.УстановитьПараметр("ОбъектРемонта", ОбъектРемонта);   
	Запрос.УстановитьПараметр("ДатаНачала", ДатаНачала);
	
	Возврат Запрос.Выполнить().Выгрузить();

КонецФункции  

#КонецОбласти