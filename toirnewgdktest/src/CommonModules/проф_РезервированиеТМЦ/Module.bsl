#Область ПрограммныйИнтерфейс

//++ Проф-ИТ, #320, Соловьев А.А., 20.11.2023

Процедура РезервированиеТМЦ() Экспорт
	
	Запрос = Новый Запрос;
	Запрос.Текст = ТекстЗапросаРезервированиеТМЦ();
		
	ТекущаяДата = ТекущаяДатаСеанса();
	
	СкладБГЛЦ = Справочники.проф_НастройкиСистемы.ПолучитьНастройкуСистемы("ОтчетОбеспечениеРемонта", 
		"СкладБГЛЦ");
	
	ПодразделенияСлужбыТОИР = Справочники.проф_НастройкиСистемы.ПолучитьСпЗначНастройкиСистемы(
		"Подразделения", "ПодразделенияСлужбыТОИР");
	
	СписокВидовИсточников = СписокВидовИсточников();
	
	Если ПолучитьФункциональнуюОпцию("торо_ИспользоватьХарактеристикиНоменклатуры") Тогда
		Запрос.Текст = СтрЗаменить(Запрос.Текст, "&Характеристика", "Регистр.Характеристика");
	Иначе
		Запрос.УстановитьПараметр("Характеристика", Справочники.ХарактеристикиНоменклатуры.ПустаяСсылка());
	КонецЕсли;
	
	Запрос.УстановитьПараметр("ТекущаяДата", ТекущаяДата);
	Запрос.УстановитьПараметр("ДатаГоризонтПланирования", КонецМесяца(ТекущаяДата));
	Запрос.УстановитьПараметр("ПериодУчетаКрайнихСроковЗаказов", КонецДня(ТекущаяДата));
	Запрос.УстановитьПараметр("СкладБГЛЦ", СкладБГЛЦ);
	Запрос.УстановитьПараметр("ПодразделенияСлужбыТОИР", ПодразделенияСлужбыТОИР);
	Запрос.УстановитьПараметр("СписокВидовИсточников", СписокВидовИсточников);
	
	Запрос.МенеджерВременныхТаблиц = Новый МенеджерВременныхТаблиц;
	РезультатЗапроса = Запрос.Выполнить();
	
	ВыборкаПоЗаказам = РезультатЗапроса.Выбрать(ОбходРезультатаЗапроса.ПоГруппировкам);
	Пока ВыборкаПоЗаказам.Следующий() Цикл
		
		ДокументОбъект = Документы.торо_РезервПодВнутреннийЗаказ.СоздатьДокумент();
		ЗаполнитьЗначенияСвойств(ДокументОбъект, ВыборкаПоЗаказам);
		ДокументОбъект.Дата = ТекущаяДата;
		
		ЗаполнитьТЧТовары(ВыборкаПоЗаказам, ДокументОбъект);
				
		Если ДокументОбъект.Товары.Количество() > 0 Тогда
			
			Попытка
				ДокументОбъект.Записать(РежимЗаписиДокумента.Проведение);
			Исключение
				ОписаниеОшибки = ПодробноеПредставлениеОшибки(ИнформацияОбОшибке());
				ЗаписьЖурналаРегистрации(
					НСтр("ru = 'Ошибка проведения при попытке резервирования из
							|регламентного задания ""Резервирование ТМЦ (Проф-ИТ)"".'"),
						УровеньЖурналаРегистрации.Ошибка,
						ДокументОбъект,,
						ОписаниеОшибки); 
			КонецПопытки;
			
		КонецЕсли;
	КонецЦикла;
	
КонецПроцедуры

//-- Проф-ИТ, #320, Соловьев А.А., 20.11.2023

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

//++ Проф-ИТ, #320, Соловьев А.А., 20.11.2023

Функция ТекстЗапросаРезервированиеТМЦ()
	
	ТекстЗапроса = 
	"ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	Регистр.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	Регистр.Номенклатура КАК Номенклатура,
	|	Регистр.Номенклатура.СхемаОбеспечения КАК СхемаОбеспечения,
	|	&Характеристика КАК Характеристика,
	|	Регистр.Склад КАК Склад,
	|	Регистр.КодСтроки КАК КодСтроки,
	|	Регистр.КОформлениюОстаток КАК КОформлениюОстаток
	|ПОМЕСТИТЬ ЗаказыНаВнутреннееПотреблениеРС
	|ИЗ
	|	РегистрНакопления.ЗаказыНаВнутреннееПотребление.Остатки КАК Регистр
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Номенклатура,
	|	Характеристика,
	|	СхемаОбеспечения,
	|	Склад,
	|	ЗаказНаВнутреннееПотребление
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	ЗаказыНаВнутреннееПотреблениеРС.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	Документ.Организация КАК Организация,
	|	Документ.Подразделение КАК Подразделение,
	|	Документ.торо_СозданПоПотребности КАК торо_СозданПоПотребности,
	|	Документ.ДокументОснование КАК ДокументОснование,
	|	Документ.проф_Назначение КАК проф_Назначение,
	|	ЗаказыНаВнутреннееПотреблениеРС.Номенклатура КАК Номенклатура,
	|	ЗаказыНаВнутреннееПотреблениеРС.СхемаОбеспечения КАК СхемаОбеспечения,
	|	ЗаказыНаВнутреннееПотреблениеРС.Характеристика КАК Характеристика,
	|	ЗаказыНаВнутреннееПотреблениеРС.Склад КАК Склад,
	|	ЗаказыНаВнутреннееПотреблениеРС.КОформлениюОстаток КАК КОформлениюОстаток,
	|	ЗаказНаВнутреннееПотреблениеТовары.проф_КЗаказу КАК проф_КЗаказу
	|ПОМЕСТИТЬ ЗаказыНаВнутреннееПотребление
	|ИЗ
	|	ЗаказыНаВнутреннееПотреблениеРС КАК ЗаказыНаВнутреннееПотреблениеРС
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.ЗаказНаВнутреннееПотребление КАК Документ
	|		ПО ЗаказыНаВнутреннееПотреблениеРС.ЗаказНаВнутреннееПотребление = Документ.Ссылка
	|			И (Документ.торо_СтатусУтверждения = ЗНАЧЕНИЕ(Перечисление.торо_СтатусыУтвержденияЗаказовНаВП.Утвержден))
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.ЗаказНаВнутреннееПотребление.Товары КАК ЗаказНаВнутреннееПотреблениеТовары
	|		ПО ЗаказыНаВнутреннееПотреблениеРС.ЗаказНаВнутреннееПотребление = ЗаказНаВнутреннееПотреблениеТовары.Ссылка
	|			И ЗаказыНаВнутреннееПотреблениеРС.Номенклатура = ЗаказНаВнутреннееПотреблениеТовары.Номенклатура
	|			И ЗаказыНаВнутреннееПотреблениеРС.Характеристика = ЗаказНаВнутреннееПотреблениеТовары.Характеристика
	|			И ЗаказыНаВнутреннееПотреблениеРС.КодСтроки = ЗаказНаВнутреннееПотреблениеТовары.КодСтроки
	|ГДЕ
	|	ЗаказыНаВнутреннееПотреблениеРС.Склад <> ЗНАЧЕНИЕ(Справочник.Склады.ПустаяСсылка)
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Номенклатура,
	|	Характеристика,
	|	СхемаОбеспечения,
	|	Склад,
	|	ЗаказНаВнутреннееПотребление
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	Регистр.Номенклатура КАК Номенклатура,
	|	Регистр.Характеристика КАК Характеристика,
	|	Регистр.Склад КАК Склад,
	|	Регистр.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	Регистр.КоличествоОстаток КАК КоличествоОстаток
	|ПОМЕСТИТЬ РезервыПоЗаказам
	|ИЗ
	|	РегистрНакопления.торо_РезервыНаСкладах.Остатки(
	|			,
	|			(Номенклатура, Характеристика) В
	|				(ВЫБРАТЬ
	|					Потребность.Номенклатура КАК Номенклатура,
	|					Потребность.Характеристика КАК Характеристика
	|				ИЗ
	|					ЗаказыНаВнутреннееПотребление КАК Потребность)) КАК Регистр
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Номенклатура,
	|	Характеристика,
	|	Склад,
	|	ЗаказНаВнутреннееПотребление
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	РезервыПоЗаказам.Номенклатура КАК Номенклатура,
	|	РезервыПоЗаказам.Характеристика КАК Характеристика,
	|	РезервыПоЗаказам.Склад КАК Склад,
	|	СУММА(РезервыПоЗаказам.КоличествоОстаток) КАК КоличествоОстаток
	|ПОМЕСТИТЬ РезервыПоНоменклатуре
	|ИЗ
	|	РезервыПоЗаказам КАК РезервыПоЗаказам
	|
	|СГРУППИРОВАТЬ ПО
	|	РезервыПоЗаказам.Номенклатура,
	|	РезервыПоЗаказам.Характеристика,
	|	РезервыПоЗаказам.Склад
	|
	|ИНДЕКСИРОВАТЬ ПО
	|	Номенклатура,
	|	Характеристика,
	|	Склад
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	проф_ЗаказыНаПеремещениеОстатки.ЗаказНаПеремещение КАК ЗаказНаПеремещение,
	|	проф_ЗаказыНаПеремещениеОстатки.Номенклатура КАК Номенклатура,
	|	проф_ЗаказыНаПеремещениеОстатки.Характеристика КАК Характеристика,
	|	СУММА(проф_ЗаказыНаПеремещениеОстатки.КОформлениюОстаток) КАК КОформлению
	|ПОМЕСТИТЬ втЗаказыНаПеремещениеОстатки
	|ИЗ
	|	РегистрНакопления.проф_ЗаказыНаПеремещение.Остатки(&ТекущаяДата, ) КАК проф_ЗаказыНаПеремещениеОстатки
	|
	|СГРУППИРОВАТЬ ПО
	|	проф_ЗаказыНаПеремещениеОстатки.Характеристика,
	|	проф_ЗаказыНаПеремещениеОстатки.Номенклатура,
	|	проф_ЗаказыНаПеремещениеОстатки.ЗаказНаПеремещение
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	втЗаказыНаПеремещениеОстатки.Номенклатура КАК Номенклатура,
	|	втЗаказыНаПеремещениеОстатки.Характеристика КАК Характеристика,
	|	СУММА(втЗаказыНаПеремещениеОстатки.КОформлению) КАК КОформлению,
	|	проф_ЗаказНаПеремещение.ДокументОснование КАК ДокументОснование
	|ПОМЕСТИТЬ втЗаказыНаПеремещение
	|ИЗ
	|	втЗаказыНаПеремещениеОстатки КАК втЗаказыНаПеремещениеОстатки
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.проф_ЗаказНаПеремещение КАК проф_ЗаказНаПеремещение
	|		ПО втЗаказыНаПеремещениеОстатки.ЗаказНаПеремещение = проф_ЗаказНаПеремещение.Ссылка
	|ГДЕ
	|	проф_ЗаказНаПеремещение.СкладОтправитель <> &СкладБГЛЦ
	|
	|СГРУППИРОВАТЬ ПО
	|	втЗаказыНаПеремещениеОстатки.Номенклатура,
	|	втЗаказыНаПеремещениеОстатки.Характеристика,
	|	проф_ЗаказНаПеремещение.ДокументОснование
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	проф_ЗапасыИПотребностиОстатки.Номенклатура КАК Номенклатура,
	|	проф_ЗапасыИПотребностиОстатки.Характеристика КАК Характеристика,
	|	проф_ЗапасыИПотребностиОстатки.Склад КАК Склад,
	|	проф_ЗапасыИПотребностиОстатки.Назначение КАК Назначение,
	|	проф_ЗапасыИПотребностиОстатки.ВНаличииОстаток - проф_ЗапасыИПотребностиОстатки.РезервироватьНаСкладеОстаток - проф_ЗапасыИПотребностиОстатки.РезервироватьПоМереПоступленияОстаток КАК ВНаличии
	|ПОМЕСТИТЬ втЗапасыИПотребностиВНаличии
	|ИЗ
	|	РегистрНакопления.проф_ЗапасыИПотребности.Остатки(
	|			&ТекущаяДата,
	|			(Номенклатура, Характеристика) В
	|				(ВЫБРАТЬ
	|					Потребность.Номенклатура КАК Номенклатура,
	|					Потребность.Характеристика КАК Характеристика
	|				ИЗ
	|					ЗаказыНаВнутреннееПотребление КАК Потребность)) КАК проф_ЗапасыИПотребностиОстатки
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	втЗапасыИПотребности.Номенклатура КАК Номенклатура,
	|	втЗапасыИПотребности.Характеристика КАК Характеристика,
	|	втЗапасыИПотребности.Назначение КАК Назначение,
	|	втЗапасыИПотребности.Склад КАК Склад,
	|	втЗапасыИПотребности.ВНаличии КАК Доступно
	|ПОМЕСТИТЬ втЗапасыИПотребностиНазначение
	|ИЗ
	|	втЗапасыИПотребностиВНаличии КАК втЗапасыИПотребности
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.проф_Назначения КАК проф_Назначения
	|		ПО втЗапасыИПотребности.Назначение = проф_Назначения.Ссылка
	|			И (проф_Назначения.Подразделение В (&ПодразделенияСлужбыТОИР))
	|
	|ОБЪЕДИНИТЬ ВСЕ
	|
	|ВЫБРАТЬ
	|	втЗапасыИПотребности.Номенклатура,
	|	втЗапасыИПотребности.Характеристика,
	|	втЗапасыИПотребности.Назначение,
	|	втЗапасыИПотребности.Склад,
	|	втЗапасыИПотребности.ВНаличии
	|ИЗ
	|	втЗапасыИПотребностиВНаличии КАК втЗапасыИПотребности
	|ГДЕ
	|	втЗапасыИПотребности.Назначение = ЗНАЧЕНИЕ(Справочник.проф_Назначения.ПустаяСсылка)
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	втЗапасыИПотребностиНазначение.Номенклатура КАК Номенклатура,
	|	втЗапасыИПотребностиНазначение.Характеристика КАК Характеристика,
	|	втЗапасыИПотребностиНазначение.Склад КАК Склад,
	|	СУММА(втЗапасыИПотребностиНазначение.Доступно) КАК Доступно
	|ПОМЕСТИТЬ втЗапасыИПотребности
	|ИЗ
	|	втЗапасыИПотребностиНазначение КАК втЗапасыИПотребностиНазначение
	|
	|СГРУППИРОВАТЬ ПО
	|	втЗапасыИПотребностиНазначение.Номенклатура,
	|	втЗапасыИПотребностиНазначение.Склад,
	|	втЗапасыИПотребностиНазначение.Характеристика
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	ЗаказыНаВнутреннееПотребление.Номенклатура КАК Номенклатура,
	|	ЗаказыНаВнутреннееПотребление.Характеристика КАК Характеристика,
	|	ЗаказыНаВнутреннееПотребление.Склад КАК Склад,
	|	ЗаказыНаВнутреннееПотребление.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	ЗаказыНаВнутреннееПотребление.проф_КЗаказу КАК КЗаказу,
	|	ЗаказыНаВнутреннееПотребление.Организация КАК Организация,
	|	ЗаказыНаВнутреннееПотребление.Подразделение КАК Подразделение,
	|	ЗаказыНаВнутреннееПотребление.проф_Назначение КАК Назначение,
	|	ЛОЖЬ КАК ДействиеВыполнено,
	|	ВЫБОР
	|		КОГДА ВЫБОР
	|				КОГДА ЕСТЬNULL(втЗапасыИПотребности.Доступно, 0) > 0
	|					ТОГДА втЗапасыИПотребности.Доступно - ЕСТЬNULL(РезервыПоНоменклатуре.КоличествоОстаток, 0)
	|				ИНАЧЕ 0
	|			КОНЕЦ < ЗаказыНаВнутреннееПотребление.КОформлениюОстаток - ЕСТЬNULL(РезервыПоЗаказам.КоличествоОстаток, 0) - ЕСТЬNULL(втЗаказыНаПеремещение.КОформлению, 0)
	|			ТОГДА ВЫБОР
	|					КОГДА ЕСТЬNULL(втЗапасыИПотребности.Доступно, 0) > 0
	|						ТОГДА втЗапасыИПотребности.Доступно - ЕСТЬNULL(РезервыПоНоменклатуре.КоличествоОстаток, 0)
	|					ИНАЧЕ 0
	|				КОНЕЦ
	|		ИНАЧЕ ЗаказыНаВнутреннееПотребление.КОформлениюОстаток - ЕСТЬNULL(РезервыПоЗаказам.КоличествоОстаток, 0) - ЕСТЬNULL(втЗаказыНаПеремещение.КОформлению, 0)
	|	КОНЕЦ КАК РекомендованоКРезерву,
	|	ВЫБОР
	|		КОГДА ЕСТЬNULL(втЗапасыИПотребности.Доступно, 0) > 0
	|			ТОГДА втЗапасыИПотребности.Доступно - ЕСТЬNULL(РезервыПоНоменклатуре.КоличествоОстаток, 0)
	|		ИНАЧЕ 0
	|	КОНЕЦ КАК Доступно,
	|	ЗаказыНаВнутреннееПотребление.Подразделение В (&ПодразделенияСлужбыТОИР) КАК ПодразделениеСлужбыТОиР
	|ПОМЕСТИТЬ втИтоговая
	|ИЗ
	|	ЗаказыНаВнутреннееПотребление КАК ЗаказыНаВнутреннееПотребление
	|		ЛЕВОЕ СОЕДИНЕНИЕ втЗапасыИПотребности КАК втЗапасыИПотребности
	|		ПО ЗаказыНаВнутреннееПотребление.Номенклатура = втЗапасыИПотребности.Номенклатура
	|			И ЗаказыНаВнутреннееПотребление.Характеристика = втЗапасыИПотребности.Характеристика
	|			И ЗаказыНаВнутреннееПотребление.Склад = втЗапасыИПотребности.Склад
	|		ЛЕВОЕ СОЕДИНЕНИЕ РезервыПоЗаказам КАК РезервыПоЗаказам
	|		ПО ЗаказыНаВнутреннееПотребление.Номенклатура = РезервыПоЗаказам.Номенклатура
	|			И ЗаказыНаВнутреннееПотребление.Характеристика = РезервыПоЗаказам.Характеристика
	|			И ЗаказыНаВнутреннееПотребление.Склад = РезервыПоЗаказам.Склад
	|			И ЗаказыНаВнутреннееПотребление.ЗаказНаВнутреннееПотребление = РезервыПоЗаказам.ЗаказНаВнутреннееПотребление
	|		ЛЕВОЕ СОЕДИНЕНИЕ РезервыПоНоменклатуре КАК РезервыПоНоменклатуре
	|		ПО ЗаказыНаВнутреннееПотребление.Номенклатура = РезервыПоНоменклатуре.Номенклатура
	|			И ЗаказыНаВнутреннееПотребление.Характеристика = РезервыПоНоменклатуре.Характеристика
	|			И ЗаказыНаВнутреннееПотребление.Склад = РезервыПоНоменклатуре.Склад
	|		ЛЕВОЕ СОЕДИНЕНИЕ втЗаказыНаПеремещениеОстатки КАК втЗаказыНаПеремещение
	|		ПО ЗаказыНаВнутреннееПотребление.Номенклатура = втЗаказыНаПеремещение.Номенклатура
	|			И ЗаказыНаВнутреннееПотребление.Характеристика = втЗаказыНаПеремещение.Характеристика
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	втИтоговая.Номенклатура КАК Номенклатура,
	|	втИтоговая.Характеристика КАК Характеристика,
	|	втИтоговая.Назначение КАК Назначение,
	|	втИтоговая.Склад КАК Склад,
	|	втИтоговая.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	втИтоговая.РекомендованоКРезерву КАК Количество,
	|	втИтоговая.Организация КАК Организация,
	|	втИтоговая.Подразделение КАК Подразделение,
	|	втИтоговая.ДействиеВыполнено КАК ДействиеВыполнено,
	|	втИтоговая.Доступно КАК Доступно,
	|	втИтоговая.ПодразделениеСлужбыТОиР КАК ПодразделениеСлужбыТОиР
	|ПОМЕСТИТЬ втТовары
	|ИЗ
	|	втИтоговая КАК втИтоговая
	|ГДЕ
	|	втИтоговая.РекомендованоКРезерву > 0
	|	И втИтоговая.КЗаказу = ЛОЖЬ
	|	И втИтоговая.Доступно > 0
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	ТоварыНаСкладахОстатки.Номенклатура КАК Номенклатура,
	|	ТоварыНаСкладахОстатки.Характеристика КАК Характеристика,
	|	ТоварыНаСкладахОстатки.Назначение КАК Назначение,
	|	ТоварыНаСкладахОстатки.Склад КАК Склад,
	|	ТоварыНаСкладахОстатки.ВНаличииОстаток КАК Доступно
	|ПОМЕСТИТЬ втДанныеРегистраТоварыНаСкладах
	|ИЗ
	|	РегистрНакопления.проф_ТоварыНаСкладах.Остатки(
	|			&ТекущаяДата,
	|			(Номенклатура, Характеристика, Склад) В
	|				(ВЫБРАТЬ
	|					втТовары.Номенклатура КАК Номенклатура,
	|					втТовары.Характеристика КАК Характеристика,
	|					втТовары.Склад КАК Склад
	|				ИЗ
	|					втТовары КАК втТовары)) КАК ТоварыНаСкладахОстатки
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ РАЗРЕШЕННЫЕ
	|	втТовары.Организация КАК Организация,
	|	втТовары.Подразделение КАК Подразделение,
	|	втТовары.Склад КАК Склад,
	|	втТовары.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	втТовары.Номенклатура КАК Номенклатура,
	|	втТовары.Назначение КАК Назначение,
	|	втТовары.Характеристика КАК Характеристика,
	|	втТовары.Количество КАК Количество,
	|	ТоварыНаСкладахОстатки.Доступно КАК Доступно,
	|	1 КАК Приоритет
	|ПОМЕСТИТЬ втТоварыДоступность
	|ИЗ
	|	втТовары КАК втТовары
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ втДанныеРегистраТоварыНаСкладах КАК ТоварыНаСкладахОстатки
	|		ПО (ТоварыНаСкладахОстатки.Номенклатура = втТовары.Номенклатура)
	|			И (ТоварыНаСкладахОстатки.Характеристика = втТовары.Характеристика)
	|			И (ТоварыНаСкладахОстатки.Склад = втТовары.Склад)
	|			И (ТоварыНаСкладахОстатки.Назначение = втТовары.Назначение)
	|			И (ТоварыНаСкладахОстатки.Доступно > 0)
	|
	|ОБЪЕДИНИТЬ ВСЕ
	|
	|ВЫБРАТЬ
	|	втТовары.Организация,
	|	втТовары.Подразделение,
	|	втТовары.Склад,
	|	втТовары.ЗаказНаВнутреннееПотребление,
	|	втТовары.Номенклатура,
	|	ЗНАЧЕНИЕ(Справочник.проф_Назначения.ПустаяСсылка),
	|	втТовары.Характеристика,
	|	втТовары.Количество,
	|	ТоварыНаСкладахОстатки.Доступно,
	|	2
	|ИЗ
	|	втТовары КАК втТовары
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ втДанныеРегистраТоварыНаСкладах КАК ТоварыНаСкладахОстатки
	|		ПО (ТоварыНаСкладахОстатки.Номенклатура = втТовары.Номенклатура)
	|			И (ТоварыНаСкладахОстатки.Характеристика = втТовары.Характеристика)
	|			И (ТоварыНаСкладахОстатки.Склад = втТовары.Склад)
	|			И (ТоварыНаСкладахОстатки.Назначение <> втТовары.Назначение)
	|			И (ТоварыНаСкладахОстатки.Назначение = ЗНАЧЕНИЕ(Справочник.проф_Назначения.ПустаяСсылка))
	|			И (ТоварыНаСкладахОстатки.Доступно > 0)
	|
	|ОБЪЕДИНИТЬ ВСЕ
	|
	|ВЫБРАТЬ
	|	втТовары.Организация,
	|	втТовары.Подразделение,
	|	втТовары.Склад,
	|	втТовары.ЗаказНаВнутреннееПотребление,
	|	втТовары.Номенклатура,
	|	ТоварыНаСкладахОстатки.Назначение,
	|	втТовары.Характеристика,
	|	втТовары.Количество,
	|	ТоварыНаСкладахОстатки.Доступно,
	|	3
	|ИЗ
	|	втТовары КАК втТовары
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ втДанныеРегистраТоварыНаСкладах КАК ТоварыНаСкладахОстатки
	|			ВНУТРЕННЕЕ СОЕДИНЕНИЕ Справочник.проф_Назначения КАК проф_Назначения
	|			ПО ТоварыНаСкладахОстатки.Назначение = проф_Назначения.Ссылка
	|				И (проф_Назначения.Подразделение В (&ПодразделенияСлужбыТОИР))
	|		ПО (ТоварыНаСкладахОстатки.Номенклатура = втТовары.Номенклатура)
	|			И (ТоварыНаСкладахОстатки.Характеристика = втТовары.Характеристика)
	|			И (ТоварыНаСкладахОстатки.Склад = втТовары.Склад)
	|			И (ТоварыНаСкладахОстатки.Назначение <> втТовары.Назначение)
	|			И (ТоварыНаСкладахОстатки.Доступно > 0)
	|;
	|
	|////////////////////////////////////////////////////////////////////////////////
	|ВЫБРАТЬ
	|	втТоварыДоступность.Организация КАК Организация,
	|	втТоварыДоступность.Подразделение КАК Подразделение,
	|	втТоварыДоступность.Склад КАК Склад,
	|	втТоварыДоступность.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	|	втТоварыДоступность.Номенклатура КАК Номенклатура,
	|	втТоварыДоступность.Назначение КАК проф_НазначениеИсходное,
	|	ДокументЗаказНаВнутреннееПотребление.проф_Назначение КАК проф_Назначение,
	|	ДокументЗаказНаВнутреннееПотребление.Ответственный КАК Ответственный,
	|	втТоварыДоступность.Характеристика КАК Характеристика,
	|	МАКСИМУМ(втТоварыДоступность.Количество) КАК Количество,
	|	СУММА(втТоварыДоступность.Доступно) КАК Доступно,
	|	втТоварыДоступность.Приоритет КАК Приоритет
	|ИЗ
	|	втТоварыДоступность КАК втТоварыДоступность
	|		ВНУТРЕННЕЕ СОЕДИНЕНИЕ Документ.ЗаказНаВнутреннееПотребление КАК ДокументЗаказНаВнутреннееПотребление
	|		ПО втТоварыДоступность.ЗаказНаВнутреннееПотребление = ДокументЗаказНаВнутреннееПотребление.Ссылка
	|
	|СГРУППИРОВАТЬ ПО
	|	втТоварыДоступность.Склад,
	|	втТоварыДоступность.ЗаказНаВнутреннееПотребление,
	|	ДокументЗаказНаВнутреннееПотребление.Ответственный,
	|	втТоварыДоступность.Назначение,
	|	втТоварыДоступность.Характеристика,
	|	втТоварыДоступность.Номенклатура,
	|	втТоварыДоступность.Подразделение,
	|	втТоварыДоступность.Организация,
	|	втТоварыДоступность.Приоритет,
	|	ДокументЗаказНаВнутреннееПотребление.проф_Назначение
	|
	|УПОРЯДОЧИТЬ ПО
	|	Приоритет,
	|	втТоварыДоступность.Назначение
	|ИТОГИ
	|	МАКСИМУМ(Организация),
	|	МАКСИМУМ(Подразделение),
	|	МАКСИМУМ(Склад),
	|	МАКСИМУМ(Ответственный),
	|	МАКСИМУМ(Количество),
	|	МИНИМУМ(Приоритет)
	|ПО
	|	ЗаказНаВнутреннееПотребление,
	|	Номенклатура
	|АВТОУПОРЯДОЧИВАНИЕ";
	
	Возврат ТекстЗапроса;
	
КонецФункции

Функция СписокВидовИсточников()
	
	СписокВидовИсточников = Новый Массив;
	
	Если ПолучитьФункциональнуюОпцию("торо_ИспользоватьВнешниеОснованияДляРабот") Тогда
		СписокВидовИсточников.Добавить(Перечисления.торо_ВидыДокументовНачалаЦепочкиРемонтов.ВнешнееОснованиеДляРабот);
	КонецЕсли;
	
	Если ПолучитьФункциональнуюОпцию("торо_УчетВыявленныхДефектовОборудования") Тогда
		СписокВидовИсточников.Добавить(Перечисления.торо_ВидыДокументовНачалаЦепочкиРемонтов.ВыявленныйДефект);
	КонецЕсли;
	
	Если ПолучитьФункциональнуюОпцию("торо_ИспользоватьППР") Тогда
		СписокВидовИсточников.Добавить(Перечисления.торо_ВидыДокументовНачалаЦепочкиРемонтов.ПланГрафикППР);
	КонецЕсли;
	
	Возврат СписокВидовИсточников;
	
КонецФункции

Процедура ЗаполнитьТЧТовары(ВыборкаПоЗаказам, ДокументОбъект)
	
	ВыборкаПоНоменклатуре = ВыборкаПоЗаказам.Выбрать(ОбходРезультатаЗапроса.ПоГруппировкам);
	Пока ВыборкаПоНоменклатуре.Следующий() Цикл
		
		КоличествоДляРаспределения = ВыборкаПоНоменклатуре.Количество;
		
		Выборка = ВыборкаПоНоменклатуре.Выбрать();
		Пока Выборка.Следующий() Цикл
			
			Если КоличествоДляРаспределения = 0 Тогда 
				Прервать;
			КонецЕсли;
			
			НоваяСтрока = ДокументОбъект.Товары.Добавить();
			ЗаполнитьЗначенияСвойств(НоваяСтрока, Выборка,
				"Номенклатура, Характеристика, проф_Назначение, проф_НазначениеИсходное");
			КоличествоРаспределено = Мин(Выборка.Доступно, КоличествоДляРаспределения);
			НоваяСтрока.КоличествоУпаковок = КоличествоРаспределено;
			НоваяСтрока.Количество = НоваяСтрока.КоличествоУпаковок;
			КоличествоДляРаспределения = КоличествоДляРаспределения - КоличествоРаспределено;
			
		КонецЦикла;
		
	КонецЦикла;
	
КонецПроцедуры

//-- Проф-ИТ, #320, Соловьев А.А., 20.11.2023

#КонецОбласти