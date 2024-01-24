#Область ОбработчикиСобытий

&ИзменениеИКонтроль("ОбработкаПроведения")
Процедура проф_ОбработкаПроведения(Отказ, РежимПроведения)

	Движения.торо_ВыявленныеДефекты.Записывать 	= Истина;
	Для Каждого ТекСтрокаСписокДефектов Из СписокДефектов Цикл
		// регистр сведений торо_ВыявленныеДефекты 
		Движение = Движения.торо_ВыявленныеДефекты.Добавить();
		ЗаполнитьЗначенияСвойств(Движение, ТекСтрокаСписокДефектов);
		Движение.Период                                 = ДатаОбнаружения;
		Движение.ДефектВыявленВПроцессе                 = ДефектВыявленВПроцессе;
		Движение.Организация                            = Организация;
		Движение.Подразделение                          = Подразделение;
	КонецЦикла;

	Если НЕ Отказ Тогда
		УстановитьУправляемыеБлокировки();
	КонецЕсли;

	ТаблицаВО = торо_УчетДефектов.СформироватьТаблицуВнешнихОснований(Ссылка);
	ДвиженияПоРегистру_торо_ВнешниеОснованияДляРабот(РежимПроведения, ТаблицаВО, Отказ);

	ДвиженияПоРегистру_торо_РецидивныеДефекты();
	ДвиженияПоРегистру_торо_ГарантийностьРемонтов();
	ДвиженияПоРегистру_торо_ВсеДокументыПоРемонтам();

	торо_РаботаСоСтатусамиДокументовСервер.УстановитьСтатусРемонтовПриПроведении(ЭтотОбъект);
	торо_РаботаСоСтатусамиДокументовСервер.УстановитьСтатусДокумента(Ссылка, Ссылка, Перечисления.торо_СтатусыДокументов.Зарегистрирован);
	торо_РаботаСоСтатусамиДокументовСервер.ИзменитьСтатусыДокументовРемонта(Ссылка);

	торо_Ремонты.ОбновитьЗаписиНезависимыхРегистровПоРемонтам(ЭтотОбъект, РежимЗаписиДокумента.Проведение);
	
	#Вставка
	//++ПрофИТ.Карпов Д.Ю.#64.24082023. При загрузке из ЕРП необходимо по требованию:
	//1 - не создавать документ состояние
	//2 - Уже после проведения дефекта создавать документ закрытия ремонтов
	Если ДополнительныеСвойства.Свойство("ТребуетсяЗакрытиеЗаявокИРемонтов") Тогда
		проф_СоздатьЗакрытиеЗаявокИРемонтов();	
	Иначе
		//++ Проф-ИТ, #222, Башинская А.Ю., 03.08.2023, создание док.Состояние ОР на основании 
		проф_СоздатьСостояниеОбъектовРемонтаНаОсновании();
		//-- Проф-ИТ, #222, Башинская А.Ю., 03.08.2023 
	КонецЕсли; 
	//--ПрофИТ.Карпов Д.Ю.#64.24082023. При загрузке из ЕРП необходимо по требованию:
	#КонецВставки 
	
КонецПроцедуры  

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

//++ Проф-ИТ, #222, Башинская А.Ю., 01.11.2023

Процедура проф_СоздатьСостояниеОбъектовРемонтаНаОсновании()
	
	//++ Проф-ИТ, #326, Соловьев А.А., 01.11.2023		
	Если ДополнительныеСвойства.Свойство("проф_ЭксплуатацияВозможна")
		И ДополнительныеСвойства.проф_ЭксплуатацияВозможна <> проф_ЭксплуатацияВозможна Тогда
		НайденныйДокумент = Неопределено;
	Иначе
		НайденныйДокумент = проф_НайтиСозданныйРанееДокументСостояниеОР();
	КонецЕсли;
	//-- Проф-ИТ, #326, Соловьев А.А., 01.11.2023
	
	Если НайденныйДокумент = Неопределено Тогда
		// Нужно создавать документ
		ДокументСОР = Документы.торо_СостоянияОбъектовРемонта.СоздатьДокумент();
		ДокументСОР.Дата = ТекущаяДатаСеанса();
	Иначе  
		// Обновим уже созданный документ
		ДокументСОР = НайденныйДокумент.ПолучитьОбъект(); 
		ДокументСОР.ОбъектыРемонта.Очистить();
	КонецЕсли;
	
	ДокументСОР.Заполнить(Ссылка);
	ДокументСОР.Автор = Ссылка.Автор;
	ДокументСОР.Ответственный = Ссылка.Ответственный;	
	ДокументСОР.Записать(РежимЗаписиДокумента.Проведение); 
	
КонецПроцедуры

Функция проф_НайтиСозданныйРанееДокументСостояниеОР()
	
	Запрос = Новый Запрос;
	Запрос.Текст = 
	"ВЫБРАТЬ
	|	торо_СостоянияОбъектовРемонта.Ссылка КАК Ссылка
	|ИЗ
	|	Документ.торо_СостоянияОбъектовРемонта КАК торо_СостоянияОбъектовРемонта
	|ГДЕ
	|	НЕ торо_СостоянияОбъектовРемонта.ПометкаУдаления   
	|	И торо_СостоянияОбъектовРемонта.Проведен
	|	И торо_СостоянияОбъектовРемонта.ДокументОснование = &ДокументОснование";
	Запрос.УстановитьПараметр("ДокументОснование", Ссылка);
	РезультатЗапроса = Запрос.Выполнить();	
	ВыборкаДетальныеЗаписи = РезультатЗапроса.Выбрать();
	
	Пока ВыборкаДетальныеЗаписи.Следующий() Цикл
		Возврат	ВыборкаДетальныеЗаписи.Ссылка;
	КонецЦикла;
	
	Возврат Неопределено;
	
КонецФункции

//-- Проф-ИТ, #222, Башинская А.Ю., 01.11.2023

//++ Проф-ИТ, #64, Карпов Д. Ю., 24.08.2023

Процедура проф_СоздатьЗакрытиеЗаявокИРемонтов() 
		
	Запрос = Новый Запрос;
	Запрос.Текст = "ВЫБРАТЬ
				|	торо_ВыявленныеДефектыСписокДефектов.Ссылка КАК Документ,
				|	торо_ВыявленныеДефектыСписокДефектов.ОбъектРемонта КАК ОбъектРемонта,
				|	ЕСТЬNULL(торо_ЗакрытиеЗаявокИРемонтовЗакрываемыеРемонты.Ссылка, ЗНАЧЕНИЕ(Документ.торо_ЗакрытиеЗаявокИРемонтов.ПустаяСсылка)) КАК ЗакрытиеЗаявок,
				|	торо_ВыявленныеДефектыСписокДефектов.ID КАК ID,
				|	торо_ВыявленныеДефекты.Дата КАК Дата,
				|	торо_ВыявленныеДефекты.Организация КАК Организация,
				|	торо_ВыявленныеДефекты.Автор КАК Автор,
				|	ЕСТЬNULL(Организации.Префикс, """") КАК ПрефиксОрганизации
				|ИЗ
				|	Документ.торо_ВыявленныеДефекты.СписокДефектов КАК торо_ВыявленныеДефектыСписокДефектов
				|		ЛЕВОЕ СОЕДИНЕНИЕ Документ.торо_ЗакрытиеЗаявокИРемонтов.ЗакрываемыеРемонты КАК торо_ЗакрытиеЗаявокИРемонтовЗакрываемыеРемонты
				|		ПО (торо_ВыявленныеДефектыСписокДефектов.Ссылка = торо_ЗакрытиеЗаявокИРемонтовЗакрываемыеРемонты.ДокументОснование)
				|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.торо_ВыявленныеДефекты КАК торо_ВыявленныеДефекты
				|			ЛЕВОЕ СОЕДИНЕНИЕ Справочник.Организации КАК Организации
				|			ПО (торо_ВыявленныеДефекты.Организация = Организации.Ссылка)
				|		ПО (торо_ВыявленныеДефектыСписокДефектов.Ссылка = торо_ВыявленныеДефекты.Ссылка)
				|ГДЕ
				|	торо_ВыявленныеДефекты.Ссылка В (&Документ)";
	
	Запрос.УстановитьПараметр("Документ", Ссылка);
	ВыявленныйДефект = Запрос.выполнить().Выбрать();
	//++ Проф-ИТ, #378, Корнилов М.С., 01.12.2023
	//ПричинаЗакрытия = Константы.торо_ПричинаЗакрытияРемонтов.Получить();
	ПричинаЗакрытия = Справочники.проф_НастройкиСистемы.ПолучитьНастройкуСистемы("ФормированиеЗакрытиеЗаявокИРемонтов", 
		"ПричинаЗакрытияРемонтов");
	//-- Проф-ИТ, #378, Корнилов М.С., 01.12.2023
	ВидРемонта 		= Константы.торо_ВидРемонтаПриВводеНаОснованииВыявленныхДефектов.Получить();
	Пока ВыявленныйДефект.Следующий() Цикл
		
		ЗакрытиеЗаявокИРемонтов = ВыявленныйДефект.ЗакрытиеЗаявок;
		
		Если НЕ ЗначениеЗаполнено(ЗакрытиеЗаявокИРемонтов) Тогда
			
			ДокОбъект = Документы.торо_ЗакрытиеЗаявокИРемонтов.СоздатьДокумент();  
			ДокОбъект.Дата             = ВыявленныйДефект.Дата+1;			 
			ДокОбъект.Автор            = ВыявленныйДефект.Автор;
			ДокОбъект.ДатаАктуальности = ВыявленныйДефект.Дата+1;
			ДокОбъект.Организация      = ВыявленныйДефект.Организация;
			ДокОбъект.Ответственный    = ВыявленныйДефект.Автор;
			
			НСтр = ДокОбъект.ЗакрываемыеРемонты.Добавить();
			НСтр.ДокументОснование = ВыявленныйДефект.Документ;
			НСтр.ОбъектРемонта     = ВыявленныйДефект.ОбъектРемонта;
			НСтр.Отменен           = Истина;
			НСтр.Перенесен         = Ложь;
			НСтр.ПричинаЗакрытия   = ПричинаЗакрытия; 
			НСтр.ID   				= ВыявленныйДефект.ID; 
			НСтр.ВидРемонта   		= ВидРемонта;
			ДокОбъект.ДополнительныеСвойства.Вставить("ИзОбработки",Истина); 
			ДокОбъект.ОбменДанными.Загрузка = Истина;
			ПрефиксНомера =  Сред(Номер,1,4);
			ДокОбъект.УстановитьНовыйНомер(ПрефиксНомера+"-");

			ДокОбъект.Записать();
			ДокОбъект.ОбменДанными.Загрузка = Ложь;
			ДокОбъект.Записать(РежимЗаписиДокумента.Проведение);
		КонецЕсли;
		
	КонецЦикла;    

КонецПроцедуры

//-- Проф-ИТ, #64, Карпов Д. Ю., 24.08.2023

//++ Проф-ИТ, #133, Антонов Б.Л., 21.09.2023, Доработка
&После("ПередЗаписью")
Процедура проф_ПередЗаписью(Отказ, РежимЗаписи, РежимПроведения)
	
	КоличествоРолейУтверждаю = 0;
	
	Для Каждого Строка Из ЭтотОбъект.проф_Подписанты Цикл
		Если Строка.Роль = ПредопределенноеЗначение("Перечисление.проф_РолиПодписантов.Утверждаю") Тогда
			КоличествоРолейУтверждаю = КоличествоРолейУтверждаю + 1;
		КонецЕсли;
	КонецЦикла; 
	
	Если КоличествоРолейУтверждаю > 1 Тогда
		Сообщение = Новый СообщениеПользователю;
		Сообщение.Текст = "В документе, в подписантах, указано более одного сотрудника с ролью ""Утверждаю""!";
		Сообщение.Сообщить(); 
		
		Отказ = Истина;
	КонецЕсли;
	
	//++ Проф-ИТ, #326, Соловьев А.А., 01.11.2023
	ДополнительныеСвойства.Вставить("проф_ЭксплуатацияВозможна", Ссылка.проф_ЭксплуатацияВозможна);
	//-- Проф-ИТ, #326, Соловьев А.А., 01.11.2023
	
	//++ Проф-ИТ, #401, Соловьев А.А., 11.12.2023
	Если РежимЗаписи = РежимЗаписиДокумента.Проведение Тогда 
		проф_ОбщегоНазначенияВызовСервера.ПроверитьПризнакПодразделенияОрганизации(ЭтотОбъект["Подразделение"], Отказ);
	КонецЕсли;
	//-- Проф-ИТ, #401, Соловьев А.А., 11.12.2023
	
КонецПроцедуры
//-- Проф-ИТ, #133, Антонов Б.Л., 21.09.2023, Доработка
#КонецОбласти
