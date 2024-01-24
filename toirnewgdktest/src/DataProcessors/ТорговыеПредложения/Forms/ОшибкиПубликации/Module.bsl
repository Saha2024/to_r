
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)
	
	Параметры.Свойство("ПрайсЛист"                        , ПрайсЛист);
	
	ПараметрПозицииТорговыхПредложений = Параметры.ПозицииТорговыхПредложений;
	Если ЗначениеЗаполнено(ПараметрПозицииТорговыхПредложений) Тогда
		ПозицииТорговыхПредложений = ПоместитьВоВременноеХранилище(ПараметрПозицииТорговыхПредложений, УникальныйИдентификатор);
		МножественныйРежим = НЕ ПараметрПозицииТорговыхПредложений.Количество() = 1;
	Иначе
		МножественныйРежим = Истина;
	КонецЕсли;
	
	АвтоНавигационнаяСсылка = Ложь;
	НавигационнаяСсылка = "e1cib/app/" + ЭтотОбъект.ИмяФормы;
	
	ЦветФонаШапкиТаблицы = ЦветаСтиля.ЦветФонаШапкиТаблицы;
	ЦветФонаФормы = ЦветаСтиля.ЦветФонаФормы;
	
	ДлительнаяОперацияОбновлениеДанных = ЗапуститьОбновлениеДанных(
		ПозицииТорговыхПредложений, ПрайсЛист, Параметры.ТребуетсяВыполнитьПроверкуДанных, УникальныйИдентификатор);
	УстановитьУсловноеОформление();
	
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	ОжидатьОбновлениеДанных();
	ОбновитьОтображениеЭлементовФормы();
КонецПроцедуры

#КонецОбласти 

#Область ОбработчикиСобытийЭлементовШапкиФормы

&НаКлиенте
Процедура ДекорацияВсеНажатие(Элемент)
	Раздел = "";
	ПерезаполнитьДерево();
	ОбновитьОтображениеЭлементовФормы();
КонецПроцедуры

&НаКлиенте
Процедура ДекорацияСведенияОТовареНажатие(Элемент)
	Раздел = "СведенияОТоваре";
	ПерезаполнитьДерево();
	ОбновитьОтображениеЭлементовФормы();
КонецПроцедуры

&НаКлиенте
Процедура ДекорацияКонтактнаяИнформацияНажатие(Элемент)
	Раздел = "КонтактнаяИнформация";
	ПерезаполнитьДерево();
	ОбновитьОтображениеЭлементовФормы();
КонецПроцедуры

&НаКлиенте
Процедура ДекорацияУчетныеДанныеНажатие(Элемент)
	Раздел = "УчетныеДанные";
	ПерезаполнитьДерево();
	ОбновитьОтображениеЭлементовФормы();
КонецПроцедуры

&НаКлиенте
Процедура ДекорацияОшибкиПубликацииНажатие(Элемент)
	Раздел = "ОшибкиПубликацииСервиса";
	ПерезаполнитьДерево();
	ОбновитьОтображениеЭлементовФормы();
КонецПроцедуры

&НаКлиенте
Процедура ДеревоОбъектовВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	
	ТекущиеДанные = Элементы.ДеревоОбъектов.ТекущиеДанные;
	
	Если ТекущиеДанные = Неопределено 
		Или ТекущиеДанные.ЭтоРаздел
		Или НЕ ТекущиеДанные.ДоступенПереход Тогда
		Возврат;
	КонецЕсли;
	
	Если ТекущиеДанные.ПолноеИмяМетаданных = "ОшибкиПубликацииСервиса" Тогда
		ПоказатьПредупреждение(, НСтр("ru = 'Ошибка при получении данных сервисом 1С:Бизнес-сеть.
                                      |Пожалуйста, попробуйте отправить данные позже.'"));
	ИначеЕсли НЕ ТекущиеДанные.ОбработкаПричиныИсключения Тогда
	
		ТекущиеДанные.ВыполненПереход = Истина;
		
		ПараметрыОткрытияФормы = Новый Структура;
		Если СтрНайти(ТекущиеДанные.ПолноеИмяМетаданных, "Форма") > 0 Тогда
			ИмяОткрываемойФормы = ТекущиеДанные.ПолноеИмяМетаданных;
			ПараметрыОткрытияФормы.Вставить("ТорговоеПредложение", ТекущиеДанные.Ссылка);
			РежимОткрытияФормы = РежимОткрытияОкнаФормы.БлокироватьОкноВладельца;
		Иначе
			ИмяОткрываемойФормы = ТекущиеДанные.ПолноеИмяМетаданных + ".ФормаОбъекта";
			ПараметрыОткрытияФормы.Вставить("Ключ", ТекущиеДанные.Ссылка);
			РежимОткрытияФормы = РежимОткрытияОкнаФормы.Независимый;
		КонецЕсли;
		
		Форма = ОткрытьФорму(
			ИмяОткрываемойФормы, 
			ПараметрыОткрытияФормы, ЭтотОбъект, , , , , 
			РежимОткрытияФормы);
		
		Если Форма = Неопределено Тогда
			Возврат;
		КонецЕсли;
		
		Для Каждого Реквизит Из ТекущиеДанные.Реквизиты Цикл
		
			СообщениеПользователю = Новый СообщениеПользователю;
			СообщениеПользователю.ИдентификаторНазначения = Форма.УникальныйИдентификатор;
			СообщениеПользователю.Поле = Реквизит.Значение;
			СообщениеПользователю.Текст = Реквизит.Представление;
			СообщениеПользователю.Сообщить();
		
		КонецЦикла;
	
	Иначе
		
		ОписаниеОповещения = Новый ОписаниеОповещения("ПослеВводаУчетныхДанных", ЭтотОбъект);
		ТорговыеПредложенияКлиентПереопределяемый.ОбработатьНажатиеПоГиперссылкеПричиныИсключенияИзПубликации(ТекущиеДанные.Ссылка,
			ПрайсЛист, ?(ПустаяСтрока(ПозицииТорговыхПредложений), Неопределено ,ПолучитьИзВременногоХранилища(ПозицииТорговыхПредложений)),
			ЭтаФорма, Элемент, ОписаниеОповещения);
			
	КонецЕсли;
	
КонецПроцедуры

#КонецОбласти 

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура ПовторитьПроверку(Команда)
	ОбновитьДанные()
КонецПроцедуры

#КонецОбласти 

#Область СлужебныеПроцедурыИФункции

#Область РаботаСФормой

&НаСервере
Процедура УстановитьУсловноеОформление()

	#Область ОтображениеГиперссылки
	
	// Оформление строк гиперссылок
	ЭлементУсловногоОформления = УсловноеОформление.Элементы.Добавить();
	ЭлементУсловногоОформления.Представление = НСтр("ru = 'Оформление строк гиперссылок'");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовИндексКартинкиБезГиперссылки");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовПредставлениеБезГиперссылки");
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.ЭтоРаздел");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.Равно;
	ЭлементОтбора.ПравоеЗначение = Ложь;
	
	ЭлементУсловногоОформления.Оформление.УстановитьЗначениеПараметра("Видимость", Ложь);
	
	// Оформление строк разделов
	ЭлементУсловногоОформления = УсловноеОформление.Элементы.Добавить();
	ЭлементУсловногоОформления.Представление = НСтр("ru = 'Оформление строк разделов'");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовИндексКартинкиГиперссылка");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовПредставлениеГиперссылка");
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.ЭтоРаздел");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.Равно;
	ЭлементОтбора.ПравоеЗначение = Истина;
	
	ЭлементУсловногоОформления.Оформление.УстановитьЗначениеПараметра("Видимость", Ложь);
	
	// Оформление отработанных гиперссылок
	ЭлементУсловногоОформления = УсловноеОформление.Элементы.Добавить();
	ЭлементУсловногоОформления.Представление = НСтр("ru = 'Оформление отработанных гиперссылок'");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовПредставлениеГиперссылка");
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.ВыполненПереход");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.Равно;
	ЭлементОтбора.ПравоеЗначение = Истина;
	
	ЭлементУсловногоОформления.Оформление.УстановитьЗначениеПараметра("ЦветТекста", ЦветаСтиля.ЦветТекстаИспользованнаяГиперссылкаБЭД);
	
	// Оформление недоступных гиперссылок
	ЭлементУсловногоОформления = УсловноеОформление.Элементы.Добавить();
	ЭлементУсловногоОформления.Представление = НСтр("ru = 'Оформление отработанных гиперссылок'");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовПредставлениеГиперссылка");
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.ДоступенПереход");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.Равно;
	ЭлементОтбора.ПравоеЗначение = Ложь;
	
	ЭлементУсловногоОформления.Оформление.УстановитьЗначениеПараметра("ЦветТекста", ЦветаСтиля.ЦветПустойГиперссылки);
	
	#КонецОбласти 

	#Область ОтображениеОшибокИПредупреждений
	
	// Оформление ошибок и предупреждений
	ЭлементУсловногоОформления = УсловноеОформление.Элементы.Добавить();
	ЭлементУсловногоОформления.Представление = НСтр("ru = 'Оформление ошибок и предупреждений'");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоОшибокПустая");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовПустоеКоличество");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовНомерКартинкиОшибкиПравая");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоОшибокПравая");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоОшибокПустаяПравая");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовПустоеКоличествоПраво");
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.КоличествоОшибок");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.НеРавно;
	ЭлементОтбора.ПравоеЗначение = 0;
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.КоличествоПредупреждений");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.НеРавно;
	ЭлементОтбора.ПравоеЗначение = 0;
	
	ЭлементУсловногоОформления.Оформление.УстановитьЗначениеПараметра("Видимость", Ложь);
	
	// Оформление ошибок
	ЭлементУсловногоОформления = УсловноеОформление.Элементы.Добавить();
	ЭлементУсловногоОформления.Представление = НСтр("ru = 'Оформление ошибок'");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовНомерКартинкиОшибки");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоОшибок");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовНомерКартинкиПредупреждения");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоПредупреждений");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоОшибокПустаяПравая");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовПустоеКоличествоПраво");
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.КоличествоОшибок");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.НеРавно;
	ЭлементОтбора.ПравоеЗначение = 0;
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.КоличествоПредупреждений");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.Равно;
	ЭлементОтбора.ПравоеЗначение = 0;
	
	ЭлементУсловногоОформления.Оформление.УстановитьЗначениеПараметра("Видимость", Ложь);
	
	// Оформление предупреждений
	ЭлементУсловногоОформления = УсловноеОформление.Элементы.Добавить();
	ЭлементУсловногоОформления.Представление = НСтр("ru = 'Оформление предупреждений'");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовНомерКартинкиОшибки");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоОшибок");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовНомерКартинкиОшибкиПравая");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоОшибокПравая");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоОшибокПустаяПравая");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовПустоеКоличествоПраво");
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.КоличествоОшибок");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.Равно;
	ЭлементОтбора.ПравоеЗначение = 0;
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.КоличествоПредупреждений");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.НеРавно;
	ЭлементОтбора.ПравоеЗначение = 0;
	
	ЭлементУсловногоОформления.Оформление.УстановитьЗначениеПараметра("Видимость", Ложь);
	
	// Оформление отсутствия ошибок и предупреждений
	ЭлементУсловногоОформления = УсловноеОформление.Элементы.Добавить();
	ЭлементУсловногоОформления.Представление = НСтр("ru = 'Оформление отсутствия ошибок и предупреждений'");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовНомерКартинкиОшибки");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоОшибок");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовНомерКартинкиОшибкиПравая");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоОшибокПравая");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовНомерКартинкиПредупреждения");
	
	Поле                         = ЭлементУсловногоОформления.Поля.Элементы.Добавить();
	Поле.Поле                    = Новый ПолеКомпоновкиДанных("ДеревоОбъектовКоличествоПредупреждений");
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.КоличествоОшибок");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.Равно;
	ЭлементОтбора.ПравоеЗначение = 0;
	
	ЭлементОтбора                = ЭлементУсловногоОформления.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ЭлементОтбора.ЛевоеЗначение  = Новый ПолеКомпоновкиДанных("ДеревоОбъектов.КоличествоПредупреждений");
	ЭлементОтбора.ВидСравнения   = ВидСравненияКомпоновкиДанных.Равно;
	ЭлементОтбора.ПравоеЗначение = 0;
	
	ЭлементУсловногоОформления.Оформление.УстановитьЗначениеПараметра("Видимость", Ложь);
	
	#КонецОбласти
	
КонецПроцедуры
 
&НаКлиенте
Процедура ОбновитьОтображениеЭлементовФормы()

	Элементы.ГруппаРазделВсе.ЦветФона                   = ?(Раздел = "", ЦветФонаШапкиТаблицы, ЦветФонаФормы);
	Элементы.ГруппаРазделСведенияОТоваре.ЦветФона       = ?(Раздел = "СведенияОТоваре", ЦветФонаШапкиТаблицы, ЦветФонаФормы);
	Элементы.ГруппаРазделКонтактнаяИнформация.ЦветФона  = ?(Раздел = "КонтактнаяИнформация", ЦветФонаШапкиТаблицы, ЦветФонаФормы);
	Элементы.ГруппаРазделУчетныеДанные.ЦветФона         = ?(Раздел = "УчетныеДанные", ЦветФонаШапкиТаблицы, ЦветФонаФормы);
	Элементы.ГруппаРазделОшибкиПубликации.ЦветФона      = ?(Раздел = "ОшибкаПубликацииСервиса", ЦветФонаШапкиТаблицы, ЦветФонаФормы);

	Элементы.ДекорацияВсе.Гиперссылка                   = НЕ Раздел = "";
	Элементы.ДекорацияСведенияОТоваре.Гиперссылка       = НЕ Раздел = "СведенияОТоваре";
	Элементы.ДекорацияКонтактнаяИнформация.Гиперссылка  = НЕ Раздел = "КонтактнаяИнформация";
	Элементы.ДекорацияУчетныеДанные.Гиперссылка         = НЕ Раздел = "УчетныеДанные";
	Элементы.ДекорацияОшибкиПубликации.Гиперссылка      = НЕ Раздел = "ОшибкаПубликацииСервиса";

КонецПроцедуры

#КонецОбласти

#Область ОбновлениеДанных

&НаСервереБезКонтекста
Функция ЗапуститьОбновлениеДанных(ПозицииТорговыхПредложений, ПрайсЛист, ТребуетсяВыполнитьПроверкуДанных, УникальныйИдентификатор)

	НаименованиеЗадания = НСтр(
		"ru = 'Формирование списка объектов для диагностики проблем обмена с 1С:Бизнес-сеть.'");
	ИмяПроцедуры        = "Обработки.ТорговыеПредложения.ПолучитьДанныеДиагностики";
	
	ПараметрыОтбора = Новый Структура;
	
	Если ЗначениеЗаполнено(ПозицииТорговыхПредложений) Тогда
		ПараметрыОтбора.Вставить("ПозицииТорговыхПредложений" , ПолучитьИзВременногоХранилища(ПозицииТорговыхПредложений));
	КонецЕсли;

	Если ЗначениеЗаполнено(ПрайсЛист) Тогда
		ПараметрыОтбора.Вставить("ПрайсЛист"                  , ПрайсЛист);
	КонецЕсли;
	
	ПараметрыПроцедуры = Новый Структура;
	ПараметрыПроцедуры.Вставить("Отбор"                            , ПараметрыОтбора);
	ПараметрыПроцедуры.Вставить("ТребуетсяВыполнитьПроверкуДанных" , ТребуетсяВыполнитьПроверкуДанных);
	
	ПараметрыВыполнения = ДлительныеОперации.ПараметрыВыполненияВФоне(УникальныйИдентификатор);
	ПараметрыВыполнения.НаименованиеФоновогоЗадания = НаименованиеЗадания;
	ПараметрыВыполнения.ОжидатьЗавершение           = 0;
	ПараметрыВыполнения.ЗапуститьВФоне              = Истина;
	
	Возврат ДлительныеОперации.ВыполнитьВФоне(ИмяПроцедуры,
		ПараметрыПроцедуры, ПараметрыВыполнения);

КонецФункции

&НаКлиенте
Процедура ОжидатьОбновлениеДанных()

	Если ДлительнаяОперацияОбновлениеДанных.Статус = "Выполнено" Тогда
		ПослеПолученияДанных(ДлительнаяОперацияОбновлениеДанных, Неопределено);
	Иначе
		
		Элементы.ГруппаСтраницы.ТекущаяСтраница = Элементы.ГруппаОбновлениеДанных;
		
		ОписаниеОповещения = Новый ОписаниеОповещения("ПослеПолученияДанных", ЭтотОбъект);
		
		ПараметрыОжидания = ДлительныеОперацииКлиент.ПараметрыОжидания(ЭтотОбъект);
		ПараметрыОжидания.ВыводитьПрогрессВыполнения      = Ложь;
		ПараметрыОжидания.ВыводитьОкноОжидания            = Ложь;
		ПараметрыОжидания.ОповещениеПользователя.Показать = Ложь;
		ПараметрыОжидания.ВыводитьСообщения               = Ложь;
		
		ДлительныеОперацииКлиент.ОжидатьЗавершение(ДлительнаяОперацияОбновлениеДанных, 
			ОписаниеОповещения, ПараметрыОжидания);
		
	КонецЕсли;

КонецПроцедуры

&НаКлиенте
Процедура ПослеПолученияДанных(РезультатВыполнения, ДополнительныеПараметры) Экспорт
	
	Если РезультатВыполнения = Неопределено Тогда
		Возврат;
	КонецЕсли;

	Элементы.ГруппаСтраницы.ТекущаяСтраница = Элементы.ГруппаДанные;
	АдресДанныхДиагностики = РезультатВыполнения.АдресРезультата;
	ПерезаполнитьДерево();
	
	Если ДеревоОбъектов.ПолучитьЭлементы().Количество() = 0 Тогда
		Элементы.ГруппаСтраницы.ТекущаяСтраница = Элементы.ГруппаДанныеЗаполненыКорректно;
	КонецЕсли;

КонецПроцедуры
 
&НаКлиенте
Процедура ОбновитьДанные()

	ДлительнаяОперацияОбновлениеДанных = ЗапуститьОбновлениеДанных(ПозицииТорговыхПредложений, ПрайсЛист, Истина, УникальныйИдентификатор);
	ОжидатьОбновлениеДанных();

КонецПроцедуры

&НаСервере
Процедура ПерезаполнитьДеревоНаСервере()

	Дерево = РеквизитФормыВЗначение("ДеревоОбъектов");
	Дерево.Строки.Очистить();
	
	ДанныеДиагностики = ПолучитьИзВременногоХранилища(АдресДанныхДиагностики);
	
	Если ДанныеДиагностики = Неопределено Тогда
		Возврат;
	КонецЕсли;
		
	Если НЕ ПустаяСтрока(Раздел) И ДанныеДиагностики.Разделы.Найти(Раздел) = Неопределено Тогда
		Раздел = "";
	КонецЕсли;
	
	Элементы.ДекорацияВсе.Заголовок = 
		СтрШаблон(НСтр("ru = 'Все (%1)'"), ДанныеДиагностики.СтатистикаДиагностики.КоличествоВсе);
	Элементы.ДекорацияСведенияОТоваре.Заголовок = 
		СтрШаблон(НСтр("ru = 'Сведения о товаре (%1)'"), ДанныеДиагностики.СтатистикаДиагностики.КоличествоСведенияОТоваре);
	Элементы.ДекорацияКонтактнаяИнформация.Заголовок = 
		СтрШаблон(НСтр("ru = 'Контактная информация (%1)'"), ДанныеДиагностики.СтатистикаДиагностики.КоличествоКонтактнаяИнформация);
	Элементы.ДекорацияУчетныеДанные.Заголовок = 
		СтрШаблон(НСтр("ru = 'Учетные данные (%1)'"), ДанныеДиагностики.СтатистикаДиагностики.КоличествоУчетныеДанные);
	Элементы.ДекорацияОшибкиПубликации.Заголовок = 
		СтрШаблон(НСтр("ru = 'Ошибки публикации (%1)'"), ДанныеДиагностики.СтатистикаДиагностики.КоличествоОшибокПубликацииСервиса);
	
	Элементы.ДекорацияУчетныеДанные.Видимость = ДанныеДиагностики.СтатистикаДиагностики.КоличествоВсе > 0;
	Элементы.ДекорацияСведенияОТоваре.Видимость = ДанныеДиагностики.СтатистикаДиагностики.КоличествоСведенияОТоваре > 0;
	Элементы.ДекорацияКонтактнаяИнформация.Видимость = ДанныеДиагностики.СтатистикаДиагностики.КоличествоКонтактнаяИнформация > 0;
	Элементы.ДекорацияУчетныеДанные.Видимость = ДанныеДиагностики.СтатистикаДиагностики.КоличествоУчетныеДанные > 0;
	Элементы.ГруппаРазделОшибкиПубликации.Видимость = ДанныеДиагностики.СтатистикаДиагностики.КоличествоОшибокПубликацииСервиса > 0;
	
	Элементы.ГруппаРазделы.Видимость = ДанныеДиагностики.Разделы.Количество() > 1;
	
	ОбрабатываемыеПричиныИсключенияИзПубликации = Новый Массив;
	ТорговыеПредложенияПереопределяемый.ОпределитьОбрабатываемыеПричиныИсключенияИзПубликации(ОбрабатываемыеПричиныИсключенияИзПубликации);
	
	Если МножественныйРежим Тогда
		Для Каждого ЭлементМетаданных Из ДанныеДиагностики.ОбъектыМетаданных Цикл
			
			НоваяСтрока = Дерево.Строки.Добавить();
			НоваяСтрока.ПолноеИмяМетаданных = ЭлементМетаданных.Значение;
			НоваяСтрока.Представление       = ЭлементМетаданных.Представление;
			НоваяСтрока.ИндексКартинки      = 0;
			НоваяСтрока.ЭтоРаздел           = Истина;
			НоваяСтрока.НомерКартинкиОшибки = -1;
			НоваяСтрока.НомерКартинкиПредупреждения = -1;
			
			ДобавитьЭлементыВКонтекст(ДанныеДиагностики, НоваяСтрока.Строки, ОбрабатываемыеПричиныИсключенияИзПубликации, 
				ЭлементМетаданных.Значение);
			
			Если НоваяСтрока.Строки.Количество() = 0 Тогда
				Дерево.Строки.Удалить(НоваяСтрока);
			Иначе
				НоваяСтрока.Представление = 
					СтрШаблон("%1 (%2)", ЭлементМетаданных.Представление, НоваяСтрока.Строки.Количество());
			КонецЕсли;
			
		КонецЦикла;
	Иначе
		ДобавитьЭлементыВКонтекст(ДанныеДиагностики, Дерево.Строки, ОбрабатываемыеПричиныИсключенияИзПубликации);
	КонецЕсли;
	
	ЗначениеВРеквизитФормы(Дерево, "ДеревоОбъектов");
	
КонецПроцедуры

&НаСервере
Функция ПредставлениеСсылки(ДанныеДиагностики, СтрокаДиагностики)

	Если МножественныйРежим Тогда
		Возврат СтрокаДиагностики.Ссылка;
	ИначеЕсли ТипЗнч(СтрокаДиагностики.Ссылка) = Тип("Строка") Тогда
		Возврат СтрокаДиагностики.Ссылка;
	Иначе
		ЭлементМетаданных = ДанныеДиагностики.ОбъектыМетаданных.НайтиПоЗначению(СтрокаДиагностики.ПолноеИмяМетаданных);
		Возврат СтрШаблон("%1 (%2)", СтрокаДиагностики.Ссылка, ЭлементМетаданных.Представление);
	КонецЕсли;

КонецФункции
 
&НаСервере
Процедура ДобавитьЭлементыВКонтекст(ДанныеДиагностики, КонтекстДобавления, ОбрабатываемыеПричиныИсключенияИзПубликации,
		ПолноеИмяМетаданных = Неопределено)

	СтруктураОтбора = Новый Структура;
	
	Если НЕ ПустаяСтрока(Раздел) Тогда
		СтруктураОтбора.Вставить("ИмяРаздела", Раздел);
	КонецЕсли;

	Если НЕ ПолноеИмяМетаданных = Неопределено Тогда
		СтруктураОтбора.Вставить("ПолноеИмяМетаданных", ПолноеИмяМетаданных);
	КонецЕсли;
	
	Если СтруктураОтбора.Количество() = 0 Тогда
		НайденныеСтроки = ДанныеДиагностики.ТаблицаДиагностики;
	Иначе
		НайденныеСтроки = ДанныеДиагностики.ТаблицаДиагностики.НайтиСтроки(СтруктураОтбора);
	КонецЕсли;
	
	ОписаниеТипаПричиныИсключения = Метаданные.ОпределяемыеТипы.ПричинаИсключенияИзПубликацииТорговыеПредложения.Тип;
	
	Для Каждого НайденнаяСтрока Из НайденныеСтроки Цикл
		
		СтруктураПоиска = Новый Структура;
		СтруктураПоиска.Вставить("Ссылка", НайденнаяСтрока.Ссылка);
		
		СтрокиДерева = КонтекстДобавления.НайтиСтроки(СтруктураПоиска);
		
		Если СтрокиДерева.Количество() = 0 Тогда
			
			ОбработкаПричиныИсключения = ОписаниеТипаПричиныИсключения.СодержитТип(ТипЗнч(НайденнаяСтрока.Ссылка));
			
			СтрокаДерева = КонтекстДобавления.Добавить();
			СтрокаДерева.Ссылка                      = НайденнаяСтрока.Ссылка;
			СтрокаДерева.ПредставлениеСсылки         = ПредставлениеСсылки(ДанныеДиагностики, НайденнаяСтрока);
			СтрокаДерева.ИндексКартинки              = 1;
			СтрокаДерева.ПолноеИмяМетаданных         = НайденнаяСтрока.ПолноеИмяМетаданных;
			СтрокаДерева.ИмяРаздела                  = НайденнаяСтрока.ИмяРаздела;
			СтрокаДерева.ДоступенПереход             = 
				Не (ОбработкаПричиныИсключения И ОбрабатываемыеПричиныИсключенияИзПубликации.Найти(НайденнаяСтрока.Ссылка) = Неопределено);
			СтрокаДерева.ОбработкаПричиныИсключения  = ОбработкаПричиныИсключения;
			СтрокаДерева.НомерКартинкиОшибки         = -1;
			СтрокаДерева.НомерКартинкиПредупреждения = -1;
			
		Иначе
		
			СтрокаДерева = СтрокиДерева[0];
		
		КонецЕсли;
		
		Если СтрокаДерева.Реквизиты.НайтиПоЗначению(НайденнаяСтрока.ПутьКДанным) = Неопределено Тогда
			СтрокаДерева.Реквизиты.Добавить(НайденнаяСтрока.ПутьКДанным, НайденнаяСтрока.ТекстСообщения);
			
			Если НайденнаяСтрока.ЯвляетсяОшибкой Тогда
				СтрокаДерева.НомерКартинкиОшибки = 0;
				СтрокаДерева.КоличествоОшибок = СтрокаДерева.КоличествоОшибок + 1;
			ИначеЕсли НайденнаяСтрока.ЯвляетсяПредупреждением Тогда
				СтрокаДерева.НомерКартинкиПредупреждения = 0;
				СтрокаДерева.КоличествоПредупреждений = СтрокаДерева.КоличествоПредупреждений + 1;
			КонецЕсли;
			
		КонецЕсли;
	
	КонецЦикла;
	
КонецПроцедуры

&НаКлиенте
Процедура ПерезаполнитьДерево()
	
	ПерезаполнитьДеревоНаСервере();
	ОбновитьОтображениеЭлементовФормы();
	
КонецПроцедуры

&НаКлиенте
Процедура ПослеВводаУчетныхДанных(Результат, ДополнительныеПараметры) Экспорт 
	ОбновитьДанные();
КонецПроцедуры

#КонецОбласти

#КонецОбласти