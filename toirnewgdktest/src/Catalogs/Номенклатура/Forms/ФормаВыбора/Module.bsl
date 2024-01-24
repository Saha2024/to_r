
#Область ОбработчикиСобытийФормы

&НаСервере
Процедура ПриСозданииНаСервере(Отказ, СтандартнаяОбработка)

	УстановитьУсловноеОформление();
	
	Если Параметры.Свойство("ОтборПоТипуНоменклатуры") Тогда 
		УстановитьОтборПоТипуНоменклатурыПоПараметрам(ЭтаФорма);
	КонецЕсли;
	
	КодФормы = "Справочник_Номенклатура_ФормаСписка";
	                  
	// Вместо ПодборТоваровСервер.ПриСозданииНаСервере(ЭтаФорма);
	
	ПодборТоваровСервер.УстановитьЗначенияПоНастройкамФормы(ЭтаФорма);
	
	ЭтаФорма.ИспользоватьСтандартныйПоискПриПодбореТоваров = Ложь;
	
	Если ПодборТоваровКлиентСервер.ЭтоФормаВыбораНоменклатуры(ЭтаФорма) Тогда 
		УстановитьОтборыПоУмолчанию();
	КонецЕсли;
	
	ПодборТоваровКлиентСервер.УстановитьТекущиеСтраницыПоВариантуПоиска(ЭтаФорма);
	
	Если ПодборТоваровКлиентСервер.ЭтоФормаВыбораНоменклатуры(ЭтаФорма) Тогда
		ПодборТоваровКлиентСервер.НазначитьКнопкуВыбораПоУмолчанию(ЭтаФорма);
	КонецЕсли;
	
	ЭтаФорма.ИнформационнаяБазаФайловая      = ОбщегоНазначения.ИнформационнаяБазаФайловая();
	ЭтаФорма.ИспользоватьПолнотекстовыйПоиск = ОбщегоНазначенияУТВызовСервера.ИспользуетсяПолнотекстовыйПоиск("ИспользоватьПолнотекстовыйПоиск");
	
	Если ЭтаФорма.ИспользоватьПолнотекстовыйПоиск Тогда
			
		ЭтаФорма.ИндексПолнотекстовогоПоискаАктуален = ПолнотекстовыйПоискСервер.ИндексПоискаАктуален();
		
	КонецЕсли;
	
	// СтандартныеПодсистемы.ПодключаемыеКоманды
	ПараметрыРазмещения = ПодключаемыеКоманды.ПараметрыРазмещения();
	ПараметрыРазмещения.КоманднаяПанель = Элементы.СписокСтандартныйПодключаемыеКоманды;
	ПараметрыРазмещения.ПрефиксГрупп = "СписокСтандартный";
	ПодключаемыеКоманды.ПриСозданииНаСервере(ЭтаФорма, ПараметрыРазмещения);
	
	ПараметрыРазмещения = ПодключаемыеКоманды.ПараметрыРазмещения();
	ПараметрыРазмещения.КоманднаяПанель = Элементы.СписокРасширенныйПодключаемыеКоманды;
	ПараметрыРазмещения.ПрефиксГрупп = "СписокРасширенный";
	ПодключаемыеКоманды.ПриСозданииНаСервере(ЭтаФорма, ПараметрыРазмещения);
	// Конец СтандартныеПодсистемы.ПодключаемыеКоманды
	
	// ЭлектронноеВзаимодействие.РаботаСНоменклатурой
	КП = ?(Элементы.СтраницыСписков.ТекущаяСтраница = Элементы.СтраницаРасширенныйПоискНоменклатура,
		Элементы.КоманднаяПанельСписокРасширенныйПоискНоменклатура, Элементы.СписокСтандартныйПоискНоменклатура.КоманднаяПанель);
	ДС = ?(Элементы.СтраницыСписков.ТекущаяСтраница = Элементы.СтраницаРасширенныйПоискНоменклатура,
		Элементы.СписокРасширенныйПоискНоменклатура, Элементы.СписокСтандартныйПоискНоменклатура);
		
	РаботаСНоменклатурой.ПриСозданииНаСервереФормаСпискаНоменклатуры(ЭтотОбъект, КП, ДС, Ложь);
	// Конец ЭлектронноеВзаимодействие.РаботаСНоменклатурой

	СоответствиеСписковДляМультиязычности = Новый Соответствие;
	СоответствиеСписковДляМультиязычности.Вставить("СписокНоменклатура", "Справочник.Номенклатура");
	СоответствиеСписковДляМультиязычности.Вставить("ВидыНоменклатуры", "Справочник.ВидыНоменклатуры");
	торо_МультиязычностьСервер.ПриСозданииНаСервереОбработкаДинамическихСписков(ЭтотОбъект, СоответствиеСписковДляМультиязычности);
	
	ЕстьПравоРедактирования = ПравоДоступа("Редактирование", Метаданные.Справочники.Номенклатура);
	
	Элементы.КоманднаяПанельСписокСтандартныйПоискНоменклатураФормаИзменитьВыделенные.Видимость = ЕстьПравоРедактирования;
	Элементы.КоманднаяПанельСписокРасширенныйПоискНоменклатураФормаИзменитьВыделенные.Видимость = ЕстьПравоРедактирования;
	
	ЕстьПравоСоздания = ПравоДоступа("Добавление", Метаданные.Справочники.Номенклатура);
	
	Элементы.КоманднаяПанельФильтры.Видимость = ЕстьПравоСоздания;
		
	МожноРедактировать = ПравоДоступа("Редактирование", Метаданные.Справочники.Номенклатура);
	Элементы.КоманднаяПанельСписокРасширенныйПоискНоменклатураФормаИзменитьВыделенные.Видимость = МожноРедактировать;
	Элементы.СписокРасширенныйПоискНоменклатураКонтекстноеМенюИзменитьВыделенные.Видимость = МожноРедактировать;
	Элементы.КоманднаяПанельСписокСтандартныйПоискНоменклатураФормаИзменитьВыделенные.Видимость = МожноРедактировать;
	Элементы.СписокСтандартныйПоискНоменклатураКонтекстноеМенюИзменитьВыделенные.Видимость = МожноРедактировать;
	Элементы.ИерархияНоменклатурыКонтекстноеМенюИзменитьВыделенные.Видимость = МожноРедактировать;
	
	ПараметрОтбораНоменклатуры = Неопределено;
	ЕстьОтборПоНоменклатуре = Параметры.Отбор.Свойство("Ссылка", ПараметрОтбораНоменклатуры);
	Если ЕстьОтборПоНоменклатуре Тогда
		МассивНоменклатуры = Новый Массив(ПараметрОтбораНоменклатуры);
		ОтборНоменклатуры.ЗагрузитьЗначения(МассивНоменклатуры);		
		Элементы.ОтобразитьВсюНоменклатуру.Видимость = Истина;
	Иначе
		ОтобразитьВсюНоменклатуру = Истина;		
	КонецЕсли;
	
	Элементы.ГруппаКоманднаяПанельНавигацииПоВидуНоменклатуры.Доступность = ИспользоватьФильтры;
КонецПроцедуры

&НаКлиенте
Процедура ПриОткрытии(Отказ)
	ДобавитьПредставлениеОтборуНоменклатуры();
КонецПроцедуры

&НаКлиенте
Процедура ПриЗакрытии(ЗавершениеРаботы)
	
	Если НЕ ЗавершениеРаботы Тогда
		СохранитьНастройкиФормыНаСервере();
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура ОбработкаОповещения(ИмяСобытия, Параметр, Источник)
		
	Если ИмяСобытия = "Запись_Номенклатура" Тогда
		Если ЗначениеЗаполнено(СтрокаПоискаНоменклатура) Тогда
			ВыполнитьПоискНоменклатуры();
		КонецЕсли;
		Если ЗначениеЗаполнено(Источник) Тогда
			Элементы[ПодборТоваровКлиентСервер.ИмяСпискаНоменклатурыПоВариантуПоиска(ЭтаФорма)].ТекущаяСтрока = Источник;
		КонецЕсли;
	КонецЕсли;
	
	Список = ?(Элементы.СтраницыСписков.ТекущаяСтраница = Элементы.СтраницаРасширенныйПоискНоменклатура,
		Элементы.СписокРасширенныйПоискНоменклатура,
		Элементы.СписокСтандартныйПоискНоменклатура);
		
	// ЭлектронноеВзаимодействие.РаботаСНоменклатурой
	Если ИмяСобытия = РаботаСНоменклатуройКлиент.ОписаниеОповещенийПодсистемы().ЗагрузкаНоменклатуры Тогда
		Если Параметр.СозданныеОбъекты.Количество() > 0 Тогда
			Список.ТекущаяСтрока = Параметр.СозданныеОбъекты[0].Номенклатура;
		КонецЕсли;
	ИначеЕсли ИмяСобытия = РаботаСНоменклатуройКлиент.ОписаниеОповещенийПодсистемы().СопоставлениеНоменклатуры Тогда
		Список.Обновить();
	КонецЕсли;
	// Конец ЭлектронноеВзаимодействие.РаботаСНоменклатурой
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовШапкиФормы

#Область ОбработчикиСобытийСтрокПоиска

&НаКлиенте
Процедура СтрокаПоискаНоменклатураПриИзменении(Элемент)
	
	ВыполнитьПоискНоменклатуры();
	
КонецПроцедуры

&НаКлиенте
Процедура СтрокаПоискаНоменклатураОчистка(Элемент, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	СтрокаПоискаНоменклатура = "";
	
	СнятьОтборПоСтрокеПоискаНоменклатурыНаСервере();
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийФлаговТочногоСоответствия

&НаКлиенте
Процедура НайтиНоменклатуруПоТочномуСоответствиюПриИзменении(Элемент)
	
	ВыполнитьПоискНоменклатуры();
	
КонецПроцедуры

#КонецОбласти

#Область Прочее

&НаКлиенте
Процедура ВидНоменклатурыПриИзменении(Элемент)
	
	ВидНоменклатурыПриИзмененииНаСервере();

КонецПроцедуры

&НаКлиенте
Процедура ВидНоменклатурыОчистка(Элемент, СтандартнаяОбработка)
	
	СтандартнаяОбработка = Ложь;
	
КонецПроцедуры

&НаКлиенте
Процедура ИспользоватьФильтрыПриИзменении(Элемент)
	ИспользоватьФильтрыПриИзмененииНаСервере();
	Элементы.ГруппаКоманднаяПанельНавигацииПоВидуНоменклатуры.Доступность = ИспользоватьФильтры;
КонецПроцедуры

&НаКлиенте
Процедура ОтборКатегорииИерархияПереключательПриИзменении(Элемент)
	
	Если ОтборКатегорииИерархияПереключатель Тогда
		ВариантНавигации = "ПоСвойствам";
	Иначе
		ВариантНавигации = "ПоИерархии";
	КонецЕсли;

	ВариантНавигацииПриИзмененииНаСервере();

КонецПроцедуры

#КонецОбласти

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыСписок

&НаКлиенте
Процедура СписокПередНачаломДобавления(Элемент, Отказ, Копирование, Родитель, Группа)
	
	Если Не Копирование И Не Группа Тогда
		
		Отказ = Истина;
		
		СтруктураПараметров = Новый Структура("Родитель, ВидНоменклатуры,  АдресТаблицыПараметров, АдресТаблицыСопостовления");
		
		Если ИспользоватьФильтры Тогда
			
			Если ВариантНавигации = "ПоСвойствам" Тогда
				
				Если ЗначениеЗаполнено(ВидНоменклатуры) Тогда
					
					СтруктураАдресовТаблиц = АдресТаблицПараметровДереваОтборовНаСервере();
					
					СтруктураПараметров.АдресТаблицыПараметров = СтруктураАдресовТаблиц.АдресТаблицыПараметров;
					СтруктураПараметров.АдресТаблицыСопостовления = СтруктураАдресовТаблиц.АдресТаблицыСопостовления;
					
					ЭтоГруппа = ПроверитьВидНоменклатуры(ВидНоменклатуры);
					Если ЭтоГруппа = Истина Тогда
						СтруктураПараметров.ВидНоменклатуры = ПредопределенноеЗначение("Справочник.ВидыНоменклатуры.ПустаяСсылка");
					Иначе	
						СтруктураПараметров.ВидНоменклатуры = ВидНоменклатуры;
					КонецЕсли;
					
				КонецЕсли;
				
			Иначе
				
				СтруктураПараметров.Родитель = ?(Элементы.ИерархияНоменклатуры.ТекущиеДанные = Неопределено, ПредопределенноеЗначение("Справочник.Номенклатура.ПустаяСсылка"),
					Элементы.ИерархияНоменклатуры.ТекущиеДанные.Ссылка);
				
			КонецЕсли;
			
		КонецЕсли;
		
		ОткрытьФорму("Справочник.Номенклатура.ФормаОбъекта", СтруктураПараметров, ЭтаФорма);
			
	ИначеЕсли Группа Тогда
		
		ТекущиеДанные = Элементы.ИерархияНоменклатуры.ТекущиеДанные;
		
		Если ТекущиеДанные <> Неопределено Тогда
			
			Отказ = Истина;
			
			СтруктураПараметров = Новый Структура("Группа", ТекущиеДанные.Ссылка);
			ОткрытьФорму("Справочник.Номенклатура.ФормаГруппы",  СтруктураПараметров, ЭтаФорма);
			
		КонецЕсли;
		
	КонецЕсли;
	
КонецПроцедуры

&НаКлиенте
Процедура СписокПриАктивизацииСтроки(Элемент)
	
	ПодборТоваровКлиент.ПриАктивизацииСтрокиСпискаНоменклатуры(ЭтаФорма);
	
	// СтандартныеПодсистемы.ПодключаемыеКоманды
	ПодключаемыеКомандыКлиент.НачатьОбновлениеКоманд(ЭтотОбъект);
	// Конец СтандартныеПодсистемы.ПодключаемыеКоманды
	
КонецПроцедуры

&НаКлиенте
Процедура ОтобразитьВсюНоменклатуруПриИзменении(Элемент)
	
	Если ОтобразитьВсюНоменклатуру Тогда
		УдалитьОтборНоменклатуры();
	Иначе
		УстановитьОтборНоменклатуры();
	КонецЕсли; 

КонецПроцедуры

&НаКлиенте
Процедура УстановитьОтборНоменклатуры()

	ЭлементОтбора = СписокНоменклатура.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));			
	ЭлементОтбора.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("Ссылка");
	ЭлементОтбора.ВидСравнения = ВидСравненияКомпоновкиДанных.ВСписке;
	ЭлементОтбора.Использование = Истина;
	ЭлементОтбора.ПравоеЗначение = ОтборНоменклатуры;
	ЭлементОтбора.Представление = "ОтборПоНоменклатуре"; 
	
КонецПроцедуры

&НаКлиенте
Процедура УдалитьОтборНоменклатуры()

	Поле = Новый ПолеКомпоновкиДанных("Ссылка");
	Для каждого ЭлементОтбора Из СписокНоменклатура.Отбор.Элементы Цикл
		Если ЭлементОтбора.ЛевоеЗначение = Поле Тогда
			СписокНоменклатура.Отбор.Элементы.Удалить(ЭлементОтбора);
			Прервать;
		КонецЕсли;
	КонецЦикла;

КонецПроцедуры

&НаКлиенте
Процедура ДобавитьПредставлениеОтборуНоменклатуры()

	Поле = Новый ПолеКомпоновкиДанных("Ссылка");
	Для каждого ЭлементОтбора Из СписокНоменклатура.Отбор.Элементы Цикл
		Если ЭлементОтбора.ЛевоеЗначение = Поле Тогда
			ЭлементОтбора.Представление = "ОтборПоНоменклатуре";
			Прервать;
		КонецЕсли;
	КонецЦикла; 
	
КонецПроцедуры
 
#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыИерархияНоменклатуры

&НаКлиенте
Процедура ИерархияНоменклатурыПриАктивизацииСтроки(Элемент)
	
	ПодборТоваровКлиент.ПриАктивизацииСтрокиИерархииНоменклатуры(ЭтаФорма);
	
КонецПроцедуры

&НаКлиенте
Процедура ИерархияНоменклатурыПриАктивизацииСтрокиОбработчикОжидания()
	
	ПодборТоваровКлиент.ОбработчикАктивизацииСтрокиИерархииНоменклатуры(ЭтаФорма);
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийЭлементовТаблицыФормыВидыНоменклатуры

&НаКлиенте
Процедура ВидыНоменклатурыВыбор(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	Если ИспользоватьФильтры = Ложь Тогда
		СтандартнаяОбработка = Ложь;
	КонецЕсли; 
КонецПроцедуры

&НаКлиенте
Процедура ВидыНоменклатурыПроверкаПеретаскивания(Элемент, ПараметрыПеретаскивания, СтандартнаяОбработка, Строка, Поле)
	Если ТипЗнч(ПараметрыПеретаскивания.Значение) <> Тип("СправочникСсылка.ВидыНоменклатуры") Тогда
	 	СтандартнаяОбработка = Ложь;
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ВидыНоменклатурыПеретаскивание(Элемент, ПараметрыПеретаскивания, СтандартнаяОбработка, Строка, Поле)
	Если ТипЗнч(ПараметрыПеретаскивания.Значение) <> Тип("СправочникСсылка.ВидыНоменклатуры") Тогда
		СтандартнаяОбработка = Ложь;
		ВидыНоменклатурыПеретаскиваниеНаСервере(ПараметрыПеретаскивания.Значение, СтандартнаяОбработка, Строка);
	КонецЕсли;
КонецПроцедуры

&НаКлиенте
Процедура ВидыНоменклатурыПриАктивизацииСтроки(Элемент)
	ПодборТоваровКлиент.ПриАктивизацииСтрокиСпискаВидыНоменклатуры(ЭтаФорма);
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиКомандФормы

&НаКлиенте
Процедура ДобавитьЭлемент(Команда)
	
	ПараметрыФормы = Новый Структура;
	Если Элементы.ВидыНоменклатуры.ТекущиеДанные <> Неопределено Тогда
		ПараметрыФормы.Вставить("ВидНоменклатуры", Элементы.ВидыНоменклатуры.ТекущиеДанные.Ссылка); 
	КонецЕсли; 
	Если Элементы.ИерархияНоменклатуры.ТекущиеДанные <> Неопределено Тогда
		ПараметрыФормы.Вставить("Родитель", Элементы.ИерархияНоменклатуры.ТекущиеДанные.ГруппаНоменклатуры); 
	КонецЕсли; 
	
	ОткрытьФорму("Справочник.Номенклатура.Форма.ФормаЭлемента", ПараметрыФормы ,ЭтаФорма);
	
КонецПроцедуры

&НаКлиенте
Процедура ИзменитьВыделенныеРасширенныйПоиск(Команда)
	
	ГрупповоеИзменениеОбъектовКлиент.ИзменитьВыделенные(Элементы.СписокРасширенныйПоискНоменклатура);
	
КонецПроцедуры

&НаКлиенте
Процедура ИзменитьВыделенныеСтандартныйПоиск(Команда)
	
	ГрупповоеИзменениеОбъектовКлиент.ИзменитьВыделенные(Элементы.СписокСтандартныйПоискНоменклатура);
	
КонецПроцедуры

&НаКлиенте
Процедура ИзменитьВыделенныеГруппы(Команда)
	
	ГрупповоеИзменениеОбъектовКлиент.ИзменитьВыделенные(Элементы.ИерархияНоменклатуры);
	
КонецПроцедуры

&НаКлиенте
Процедура Поиск(Команда)
	
	ТекущееСообщениеПользователю = "";
	
КонецПроцедуры

&НаКлиенте
Процедура СброситьОтборыПоСвойствам(Команда)
	
	СброситьОтборыПоСвойствамНаСервере();
	
КонецПроцедуры

// СтандартныеПодсистемы.ПодключаемыеКоманды
&НаКлиенте
Процедура Подключаемый_ВыполнитьКоманду(Команда)
	
	СписокДляПечати = ?(Элементы.СтраницыСписков.ТекущаяСтраница = Элементы.СтраницаРасширенныйПоискНоменклатура,
		Элементы.СписокРасширенныйПоискНоменклатура,
		Элементы.СписокСтандартныйПоискНоменклатура);
	
	ПодключаемыеКомандыКлиент.НачатьВыполнениеКоманды(ЭтотОбъект, Команда, СписокДляПечати);
	
КонецПроцедуры
 
&НаКлиенте
Процедура Подключаемый_ПродолжитьВыполнениеКомандыНаСервере(ПараметрыВыполнения, ДополнительныеПараметры) Экспорт
	ВыполнитьКомандуНаСервере(ПараметрыВыполнения);
КонецПроцедуры
 
&НаСервере
Процедура ВыполнитьКомандуНаСервере(ПараметрыВыполнения)
	
	СписокДляПечати = ?(Элементы.СтраницыСписков.ТекущаяСтраница = Элементы.СтраницаРасширенныйПоискНоменклатура,
		Элементы.СписокРасширенныйПоискНоменклатура,
		Элементы.СписокСтандартныйПоискНоменклатура);
		
	ПодключаемыеКоманды.ВыполнитьКоманду(ЭтотОбъект, ПараметрыВыполнения, СписокДляПечати);
	
КонецПроцедуры
 
&НаКлиенте
Процедура Подключаемый_ОбновитьКоманды()
	
	СписокДляПечати = ?(Элементы.СтраницыСписков.ТекущаяСтраница = Элементы.СтраницаРасширенныйПоискНоменклатура,
		Элементы.СписокРасширенныйПоискНоменклатура,
		Элементы.СписокСтандартныйПоискНоменклатура);
		
	ПодключаемыеКомандыКлиентСервер.ОбновитьКоманды(ЭтотОбъект, СписокДляПечати);
	
КонецПроцедуры
// Конец СтандартныеПодсистемы.ПодключаемыеКоманды

// ЭлектронноеВзаимодействие.РаботаСНоменклатурой
&НаКлиенте
Процедура Подключаемый_ВыполнитьКомандуРаботаСНоменклатурой(Команда)
	РаботаСНоменклатуройКлиент.ВыполнитьПодключаемуюКоманду(ЭтотОбъект, Команда);
КонецПроцедуры

&НаКлиенте
Процедура Подключаемый_ВыборРаботаСНоменклатурой(Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка)
	РаботаСНоменклатуройКлиент.ВыборВТаблицеФормы(ЭтотОбъект, Элемент, ВыбраннаяСтрока, Поле, СтандартнаяОбработка);
КонецПроцедуры
// Конец ЭлектронноеВзаимодействие.РаботаСНоменклатурой
#КонецОбласти

#Область СлужебныеПроцедурыИФункции

&НаСервере
Процедура УстановитьОтборыПоУмолчанию()
	
	Для Каждого Элемент Из ЭтаФорма.Параметры.Отбор Цикл
		
		Если Элемент.Ключ = "ТипНоменклатуры" Тогда
			ПодборТоваровСервер.УстановитьОтборПоТипуНоменклатурыПоПараметрам(ЭтаФорма);
			Продолжить;
		КонецЕсли;
		
		Если Элемент.Ключ = "ВидНоменклатуры" Тогда
			ПодборТоваровСервер.УстановитьОтборПоВидНоменклатурыПоПараметрам(ЭтаФорма);
			Продолжить;
		КонецЕсли;
		
		Если Элемент.Ключ = "Владелец" Тогда
			Продолжить;
		КонецЕсли;
		
		Если ТипЗнч(Элемент.Значение) = Тип("ФиксированныйМассив") Тогда
			ПодборТоваровСервер.ДобавитьЭлементОтбора(ОбщегоНазначенияУТКлиентСервер.ПолучитьОтборДинамическогоСписка(ЭтаФорма.СписокНоменклатура), 
				Элемент.Ключ, Элемент.Значение, ВидСравненияКомпоновкиДанных.ВСписке, "ОтборПоУмолчанию");
		Иначе
			ПодборТоваровСервер.ДобавитьЭлементОтбора(ОбщегоНазначенияУТКлиентСервер.ПолучитьОтборДинамическогоСписка(ЭтаФорма.СписокНоменклатура), 
				Элемент.Ключ, Элемент.Значение, ВидСравненияКомпоновкиДанных.Равно, "ОтборПоУмолчанию");
		КонецЕсли;
			
	КонецЦикла;

КонецПроцедуры

&НаСервере
Процедура УстановитьУсловноеОформление()

	Элемент = УсловноеОформление.Элементы.Добавить();

	ПолеЭлемента = Элемент.Поля.Элементы.Добавить();
	ПолеЭлемента.Поле = Новый ПолеКомпоновкиДанных(Элементы.СтрокаПоискаНоменклатура.Имя);

	ОтборЭлемента = Элемент.Отбор.Элементы.Добавить(Тип("ЭлементОтбораКомпоновкиДанных"));
	ОтборЭлемента.ЛевоеЗначение = Новый ПолеКомпоновкиДанных("ПоискНоменклатурыНеУдачный");
	ОтборЭлемента.ВидСравнения = ВидСравненияКомпоновкиДанных.Равно;
	ОтборЭлемента.ПравоеЗначение = Истина;

	Элемент.Оформление.УстановитьЗначениеПараметра("ЦветФона", Новый Цвет(255, 200, 200));

КонецПроцедуры

#Область Поиск

&НаКлиенте
Процедура ВыполнитьПоискНоменклатуры()
	
	ОписаниеОповещенияПриУспехе = Новый ОписаниеОповещения("ВыполнитьПоискНоменклатурыЗвершение", ЭтотОбъект);
	ПодборТоваровКлиент.ВыполнениеРасширенногоПоискаВозможно(ЭтаФорма, ОписаниеОповещенияПриУспехе);
		
КонецПроцедуры

&НаКлиенте
Процедура ВыполнитьПоискНоменклатурыЗвершение(Результат, ДополнительныеПараметры) Экспорт 
	
	ВыполнитьПоискНоменклатурыНаСервере();
	
	ПодборТоваровКлиент.ПослеВыполненияПоискаНоменклатуры(ЭтаФорма);

КонецПроцедуры

&НаСервере
Процедура ВыполнитьПоискНоменклатурыНаСервере()
	
	ПодборТоваровСервер.ВыполнитьПоискНоменклатуры(ЭтаФорма);
	
КонецПроцедуры

&НаСервере
Процедура СнятьОтборПоСтрокеПоискаНоменклатурыНаСервере()
	
	ПодборТоваровКлиентСервер.СнятьОтборПоСтрокеПоискаНоменклатуры(ЭтаФорма);
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиСобытийНаСервере

&НаСервере
Процедура ИспользоватьФильтрыПриИзмененииНаСервере()
	
	ПодборТоваровСервер.ПриИзмененииИспользованияФильтров(ЭтаФорма);
	
КонецПроцедуры

&НаСервере
Процедура ВариантНавигацииПриИзмененииНаСервере()
	
	ПодборТоваровСервер.ПриИзмененииВариантаНавигации(ЭтаФорма);
		
КонецПроцедуры

&НаСервере
Процедура ВидНоменклатурыПриИзмененииНаСервере()
	
	ПодборТоваровСервер.ПриИзмененииВидаНоменклатуры(ЭтаФорма);
	
КонецПроцедуры

#КонецОбласти

#Область ГорячиеКлавиши

&НаКлиенте
Процедура УстановитьТекущийЭлементСтрокаПоиска(Команда)
	
	ПодборТоваровКлиент.УстановитьТекущийЭлементСтрокаПоиска(ЭтаФорма);
	
КонецПроцедуры

#КонецОбласти

#Область ОбработчикиОжидания

&НаКлиенте
Процедура СписокПриАктивизацииСтрокиОбработчикОжидания()
	
	ПодборТоваровКлиент.УстановитьТекущуюСтрокуИерархииНоменклатуры(ЭтаФорма);
	
КонецПроцедуры

#КонецОбласти

#Область Прочее

&НаСервере
Функция АдресТаблицПараметровДереваОтборовНаСервере()
	
	Структура = Новый Структура("АдресТаблицыПараметров, АдресТаблицыСопостовления");
	
	Структура.АдресТаблицыПараметров = ПодборТоваровСервер.АдресТаблицыПараметровДереваОтборов(ЭтаФорма);
	Структура.АдресТаблицыСопостовления = ПодборТоваровСервер.АдресТаблицыСопостовленияДереваОтборов(ЭтаФорма);

	Возврат Структура;
	
КонецФункции

&НаСервере
Процедура СохранитьНастройкиФормыНаСервере()
	
	ПодборТоваровСервер.СохранитьНастройкиФормы(ЭтаФорма);
	
КонецПроцедуры

&НаСервере
Процедура СброситьОтборыПоСвойствамНаСервере()
	
	ПодборТоваровСервер.СброситьОтборыПоСвойствам(ЭтаФорма);
	
КонецПроцедуры

&НаКлиенте
Процедура ВидыНоменклатурыПриАктивизацииСтрокиОбработчикОжидания()
	
	ТекущиеДанные = Элементы.ВидыНоменклатуры.ТекущиеДанные;
	
	Если ТекущиеДанные <> Неопределено Тогда
		ВидНоменклатурыПриИзмененииНаСервере();
	КонецЕсли;
	
КонецПроцедуры

&НаСервере
Процедура ВидыНоменклатурыПеретаскиваниеНаСервере(МассивНоменклатур, СтандартнаяОбработка, ВидНоменклатуры)
	
	ОбновитьСписок = Ложь;
	
	Если ОбщегоНазначения.ЗначениеРеквизитаОбъекта(ВидНоменклатуры, "ЭтоГруппа") Тогда
		СтандартнаяОбработка = Ложь;
	Иначе
		Для Каждого Номенклатура Из МассивНоменклатур Цикл
			НоменклатураОбъект = Номенклатура.ПолучитьОбъект();	
			Если НоменклатураОбъект.ВидНоменклатуры <> ВидНоменклатуры Тогда
				НоменклатураОбъект.ВидНоменклатуры = ВидНоменклатуры;
				НоменклатураОбъект.Записать();
				ОбновитьСписок = Истина;
			КонецЕсли;
		КонецЦикла;
	КонецЕсли;
	
	Если ОбновитьСписок Тогда
		Элементы[ПодборТоваровКлиентСервер.ИмяСпискаНоменклатурыПоВариантуПоиска(ЭтаФорма)].Обновить();
	КонецЕсли;
	
КонецПроцедуры

Процедура УстановитьОтборПоТипуНоменклатурыПоПараметрам(Форма)
	
	Если Форма.Параметры.ОтборПоТипуНоменклатуры = Неопределено Тогда
		Возврат;
	Иначе
		ОтборПоТипуНоменклатурыИзПараметров = Форма.Параметры.ОтборПоТипуНоменклатуры;
	КонецЕсли;
	
	Если ТипЗнч(ОтборПоТипуНоменклатурыИзПараметров) = Тип("ФиксированныйМассив") Тогда
		ОтборПоТипуНоменклатуры = Новый СписокЗначений;
		Для Каждого Значение Из ОтборПоТипуНоменклатурыИзПараметров Цикл
			ОтборПоТипуНоменклатуры.Добавить(Значение);
		КонецЦикла;
	КонецЕсли;
	
	ОбщегоНазначенияКлиентСервер.УстановитьЭлементОтбораДинамическогоСписка(
	Форма.СписокНоменклатура, "ТипНоменклатуры", ОтборПоТипуНоменклатуры, 
	ВидСравненияКомпоновкиДанных.ВСписке, "ОтборПоТипуНоменклатуры", (ОтборПоТипуНоменклатуры.Количество() > 0));
	
КонецПроцедуры

&НаСервереБезКонтекста
Функция ПроверитьВидНоменклатуры(ВидНоменклатуры)
	
	Если ВидНоменклатуры.ЭтоГруппа Тогда
		Возврат Истина;
	КонецЕсли;
		
КонецФункции


#КонецОбласти

#КонецОбласти