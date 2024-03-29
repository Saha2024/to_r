#Если Сервер Или ТолстыйКлиентОбычноеПриложение Или ВнешнееСоединение Тогда

#Область СлужебныйПрограммныйИнтерфейс

Функция ОбязателенВводКомментария(ЗаказНаВП, КодСтрокиНоменклатуры = 0) Экспорт
	
	Запрос = Новый Запрос();
	Запрос.Текст = "ВЫБРАТЬ
	               |	торо_КомментарииКУтверждениюЗаказовНаВПСрезПоследних.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	               |	торо_КомментарииКУтверждениюЗаказовНаВПСрезПоследних.Пользователь КАК Пользователь,
	               |	торо_КомментарииКУтверждениюЗаказовНаВПСрезПоследних.КодСтрокиНоменклатуры КАК КодСтрокиНоменклатуры,
	               |	торо_КомментарииКУтверждениюЗаказовНаВПСрезПоследних.Период КАК Период
	               |ПОМЕСТИТЬ ВТ_АвторыКомментариевСрезПоследних
	               |ИЗ
	               |	РегистрСведений.торо_КомментарииКУтверждениюЗаказовНаВП.СрезПоследних(
	               |			,
	               |			ЗаказНаВнутреннееПотребление = &ЗаказНаВП
	               |				И КодСтрокиНоменклатуры = &КодСтрокиНоменклатуры) КАК торо_КомментарииКУтверждениюЗаказовНаВПСрезПоследних
	               |
	               |ИНДЕКСИРОВАТЬ ПО
	               |	Период,
	               |	ЗаказНаВнутреннееПотребление,
	               |	КодСтрокиНоменклатуры
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_АвторыКомментариевСрезПоследних.ЗаказНаВнутреннееПотребление КАК ЗаказНаВнутреннееПотребление,
	               |	ВТ_АвторыКомментариевСрезПоследних.КодСтрокиНоменклатуры КАК КодСтрокиНоменклатуры,
	               |	МАКСИМУМ(ВТ_АвторыКомментариевСрезПоследних.Период) КАК Период
	               |ПОМЕСТИТЬ ВТ_ПериодыПоследнихАвтороКомментариев
	               |ИЗ
	               |	ВТ_АвторыКомментариевСрезПоследних КАК ВТ_АвторыКомментариевСрезПоследних
	               |
	               |СГРУППИРОВАТЬ ПО
	               |	ВТ_АвторыКомментариевСрезПоследних.ЗаказНаВнутреннееПотребление,
	               |	ВТ_АвторыКомментариевСрезПоследних.КодСтрокиНоменклатуры
	               |
	               |ИНДЕКСИРОВАТЬ ПО
	               |	Период,
	               |	ЗаказНаВнутреннееПотребление,
	               |	КодСтрокиНоменклатуры
	               |;
	               |
	               |////////////////////////////////////////////////////////////////////////////////
	               |ВЫБРАТЬ
	               |	ВТ_АвторыКомментариевСрезПоследних.Пользователь КАК Пользователь
	               |ИЗ
	               |	ВТ_АвторыКомментариевСрезПоследних КАК ВТ_АвторыКомментариевСрезПоследних
	               |		ВНУТРЕННЕЕ СОЕДИНЕНИЕ ВТ_ПериодыПоследнихАвтороКомментариев КАК ВТ_ПериодыПоследнихАвтороКомментариев
	               |		ПО ВТ_АвторыКомментариевСрезПоследних.Период = ВТ_ПериодыПоследнихАвтороКомментариев.Период
	               |			И ВТ_АвторыКомментариевСрезПоследних.ЗаказНаВнутреннееПотребление = ВТ_ПериодыПоследнихАвтороКомментариев.ЗаказНаВнутреннееПотребление
	               |			И ВТ_АвторыКомментариевСрезПоследних.КодСтрокиНоменклатуры = ВТ_ПериодыПоследнихАвтороКомментариев.КодСтрокиНоменклатуры";
	
	Запрос.УстановитьПараметр("ЗаказНаВП", ЗаказНаВП);
	Запрос.УстановитьПараметр("КодСтрокиНоменклатуры", КодСтрокиНоменклатуры);
	
	РезультатЗапроса = Запрос.Выполнить();
	Если РезультатЗапроса.Пустой() Тогда
	    Возврат Истина;
	КонецЕсли;
	
	ВыборкаПользователя = РезультатЗапроса.Выбрать();
	ВыборкаПользователя.Следующий();
	
	Если ВыборкаПользователя.Пользователь = Пользователи.ТекущийПользователь() Тогда
	    Возврат Ложь;
	Иначе
		Возврат Истина;
	КонецЕсли;

КонецФункции

Функция СохранитьКомментарийКУтверждению(ЗаказНаВП, Комментарий, КодСтрокиНоменклатуры = 0) Экспорт
	
	ТекущийПользователь = Пользователи.ТекущийПользователь();
	
	ТекущаяДата = ТекущаяДатаСеанса();
	
	НаборЗаписей = РегистрыСведений.торо_КомментарииКУтверждениюЗаказовНаВП.СоздатьНаборЗаписей();
	НаборЗаписей.Отбор.Период.Установить(ТекущаяДата);
	НаборЗаписей.Отбор.ЗаказНаВнутреннееПотребление.Установить(ЗаказНаВП);
	НаборЗаписей.Отбор.Пользователь.Установить(ТекущийПользователь);
	НаборЗаписей.Отбор.КодСтрокиНоменклатуры.Установить(КодСтрокиНоменклатуры);
	
	НоваяСтрокаНЗ = НаборЗаписей.Добавить();
	НоваяСтрокаНЗ.Период = ТекущаяДата;
	НоваяСтрокаНЗ.ЗаказНаВнутреннееПотребление = ЗаказНаВП;
	НоваяСтрокаНЗ.Пользователь = ТекущийПользователь;
	НоваяСтрокаНЗ.КодСтрокиНоменклатуры = КодСтрокиНоменклатуры;
	НоваяСтрокаНЗ.Комментарий = Комментарий;
	
	НаборЗаписей.Записать(Истина);

КонецФункции

#КонецОбласти

#Область СлужебныеПроцедурыИФункции

#КонецОбласти

#КонецЕсли