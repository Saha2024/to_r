//++ Проф-ИТ, #44, Соловьев А.А., 21.12.2023

#Область ОбработчикиСобытийФормы

&НаСервере
Процедура проф_ПриСозданииНаСервереПеред(Отказ, СтандартнаяОбработка)
	
	ИмяСписка = "СписокСпособовОбеспеченияСклада";
	Заголовок = НСтр("ru = 'Тип обеспечения'");
	ИмяКолонки = "ТипОбеспечения"; 
	
	НоваяКолонкаТаблицы = Элементы.Добавить(ИмяСписка + ИмяКолонки, Тип("ПолеФормы"), Элементы.СписокСпособовОбеспеченияСклада);
	НоваяКолонкаТаблицы.Заголовок = Заголовок; 
	НоваяКолонкаТаблицы.ПутьКДанным = ИмяСписка + "." + ИмяКолонки; 
	НоваяКолонкаТаблицы.Ширина = 20;
	
	Элементы.Переместить(НоваяКолонкаТаблицы, Элементы.СписокСкладов, Элементы.СписокСпособовОбеспеченияСкладаИсточникОбеспечения);
	
	ИмяСписка = "СписокСкладов";
	
	НоваяКолонкаТаблицы = Элементы.Добавить(ИмяСписка + ИмяКолонки, Тип("ПолеФормы"), Элементы.СписокСкладов);
	НоваяКолонкаТаблицы.Заголовок = Заголовок; 
	НоваяКолонкаТаблицы.ПутьКДанным = ИмяСписка + "." + ИмяКолонки; 
	НоваяКолонкаТаблицы.Ширина = 20;
	
	Элементы.Переместить(НоваяКолонкаТаблицы, Элементы.СписокСкладов, Элементы.СписокСкладовИсточникОбеспечения);
	
КонецПроцедуры

&НаСервере
&ИзменениеИКонтроль("УстановитьУсловноеОформление")
Процедура проф_УстановитьУсловноеОформление()

	ОформляемыеПоля = Новый Массив;

	#Область СписокСпособовОбеспеченияСклада
	УсловноеОформлениеСписка = СписокСпособовОбеспеченияСклада.УсловноеОформление;

	Элемент = ДобавитьЭлементОформленияСУсловием(УсловноеОформлениеСписка, "Характеристика",
	"ХарактеристикиИспользуются", Ложь);
	Элемент.Оформление.УстановитьЗначениеПараметра("Текст",			НСтр("ru = '<характеристики не используются>'"));
	Элемент.Оформление.УстановитьЗначениеПараметра("ЦветТекста",	WebЦвета.НейтральноСерый);

	Элемент = ДобавитьЭлементОформленияСУсловием(УсловноеОформлениеСписка, "СпособОбеспечения",
	"ЭтоИндивидуальнаяНастройка", Ложь);
	Элемент.Оформление.УстановитьЗначениеПараметра("Отображать",		Ложь);

	Элемент = ДобавитьЭлементОформленияСУсловием(УсловноеОформлениеСписка, "СхемаОбеспечения",
	"ЭтоИндивидуальнаяНастройка", Истина);
	Элемент.Оформление.УстановитьЗначениеПараметра("Отображать",		Ложь);

	ОформляемыеПоля.Добавить("СпособОбеспечения");
	ОформляемыеПоля.Добавить("ИсточникОбеспечения");
	ОформляемыеПоля.Добавить("СрокПоставки");
	ОформляемыеПоля.Добавить("ОбеспечиваемыйПериод");
	#Вставка
	//++ Проф-ИТ, #44, Соловьев А.А., 21.12.2023
	ОформляемыеПоля.Добавить("ТипОбеспечения");
	//-- Проф-ИТ, #44, Соловьев А.А., 21.12.2023
	#КонецВставки
	Элемент = ДобавитьЭлементОформленияСУсловием(УсловноеОформлениеСписка, ОформляемыеПоля, 
	"СпособОбеспечения",, ВидСравненияКомпоновкиДанных.НеЗаполнено);
	Элемент.Оформление.УстановитьЗначениеПараметра("Текст",			НСтр("ru = '<не установлен>'"));
	Элемент.Оформление.УстановитьЗначениеПараметра("ЦветТекста",	WebЦвета.НейтральноСерый);
	ОформляемыеПоля.Очистить();

	ОформитьИсточникОбеспечения(УсловноеОформлениеСписка, "ИсточникОбеспечения",
	"ИсточникОбеспечения", "ТипОбеспечения");

	Элемент = ДобавитьЭлементОформленияСУсловием(УсловноеОформлениеСписка, "СрокПоставки", "СрокПоставки", 0);
	ОтборЭлемента = Элемент.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ОтборЭлемента.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("СпособОбеспечения");
	ОтборЭлемента.ВидСравнения = ВидСравненияКомпоновкиДанных.Заполнено;
	Элемент.Оформление.УстановитьЗначениеПараметра("Текст",			НСтр("ru = '<в любой день>'"));

	Элемент = ДобавитьЭлементОформленияСУсловием(УсловноеОформлениеСписка, "СхемаОбеспечения",
	"СхемаОбеспечения",, ВидСравненияКомпоновкиДанных.НеЗаполнено);
	ОтборЭлемента = Элемент.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ОтборЭлемента.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("СпособОбеспечения");
	ОтборЭлемента.ВидСравнения = ВидСравненияКомпоновкиДанных.НеЗаполнено;
	Элемент.Оформление.УстановитьЗначениеПараметра("Текст",			НСтр("ru = '<не указана>'"));
	Элемент.Оформление.УстановитьЗначениеПараметра("ЦветТекста",	WebЦвета.НейтральноСерый);
	#КонецОбласти

	#Область СписокСкладов
	УсловноеОформлениеСписка = СписокСкладов.УсловноеОформление;

	ОформляемыеПоля.Добавить("ЕстьНастройкаДляТовара");
	ОформляемыеПоля.Добавить("СпособОбеспечения");
	ОформляемыеПоля.Добавить("ИсточникОбеспечения");
	ОформляемыеПоля.Добавить("СрокПоставки");
	#Вставка
	//++ Проф-ИТ, #44, Соловьев А.А., 21.12.2023
	ОформляемыеПоля.Добавить("ТипОбеспечения");
	//-- Проф-ИТ, #44, Соловьев А.А., 21.12.2023
	#КонецВставки
	Элемент = ДобавитьЭлементОформленияСУсловием(УсловноеОформлениеСписка, ОформляемыеПоля, "ЭтоГруппа", Истина);
	Элемент.Оформление.УстановитьЗначениеПараметра("Текст",			НСтр("ru = '<только для элементов>'"));
	Элемент.Оформление.УстановитьЗначениеПараметра("ЦветТекста",	WebЦвета.НейтральноСерый);
	ОформляемыеПоля.Очистить();

	Элемент = ДобавитьЭлементОформленияСУсловием(УсловноеОформлениеСписка, "ЕстьНастройкаДляТовара",
	"ЕстьНастройкаДляТовара", Ложь);
	УстановитьОтборОформеленияНеДляГруппы(Элемент);
	Элемент.Оформление.УстановитьЗначениеПараметра("Текст",		НСтр("ru = 'Схема обеспечения'"));

	Элемент = ДобавитьЭлементОформленияСУсловием(УсловноеОформлениеСписка, "ЕстьНастройкаДляТовара",
	"ЕстьНастройкаДляТовара", Истина);
	УстановитьОтборОформеленияНеДляГруппы(Элемент);
	Элемент.Оформление.УстановитьЗначениеПараметра("Текст",		НСтр("ru = 'Способ обеспечения'"));

	ОформляемыеПоля.Добавить("СпособОбеспечения");
	ОформляемыеПоля.Добавить("ИсточникОбеспечения");
	ОформляемыеПоля.Добавить("СрокПоставки");
	#Вставка
	//++ Проф-ИТ, #44, Соловьев А.А., 21.12.2023
	ОформляемыеПоля.Добавить("ТипОбеспечения");
	//-- Проф-ИТ, #44, Соловьев А.А., 21.12.2023
	#КонецВставки
	Элемент = ДобавитьЭлементОформленияСУсловием(УсловноеОформлениеСписка, ОформляемыеПоля, "СпособОбеспечения",,
	ВидСравненияКомпоновкиДанных.НеЗаполнено);
	УстановитьОтборОформеленияНеДляГруппы(Элемент);
	Элемент.Оформление.УстановитьЗначениеПараметра("Текст",			НСтр("ru = '<не установлен>'"));
	Элемент.Оформление.УстановитьЗначениеПараметра("ЦветТекста",	WebЦвета.НейтральноСерый);
	ОформляемыеПоля.Очистить();

	ОформитьИсточникОбеспечения(УсловноеОформлениеСписка, "ИсточникОбеспечения", "ИсточникОбеспечения",
	"ТипОбеспечения");

	Элемент = ДобавитьЭлементОформленияСУсловием(УсловноеОформлениеСписка, "СрокПоставки", "СрокПоставки", 0);
	УстановитьОтборОформеленияНеДляГруппы(Элемент);
	ОтборЭлемента = Элемент.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ОтборЭлемента.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("СпособОбеспечения");
	ОтборЭлемента.ВидСравнения = ВидСравненияКомпоновкиДанных.Заполнено;
	Элемент.Оформление.УстановитьЗначениеПараметра("Текст",			НСтр("ru = '<в любой день>'"));
	#КонецОбласти

КонецПроцедуры

#КонецОбласти

//-- Проф-ИТ, #44, Соловьев А.А., 21.12.2023