function colorRowsByColumn() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // Определяем текущую дату
  const now = new Date();
  const currentMonth = now.getMonth(); // Текущий месяц
  const currentYear = now.getFullYear(); // Текущий год
  const monthNames = [
    "Январь",
    "Февраль",
    "Март",
    "Апрель",
    "Май",
    "Июнь",
    "Июль",
    "Август",
    "Сентябрь",
    "Октябрь",
    "Ноябрь",
    "Декабрь",
  ];

  // Формируем название текущего листа
  const expectedSheetName = `Факт. ${monthNames[currentMonth]}. ${currentYear}`;

  // Проверяем, соответствует ли название активного листа ожидаемому
  if (sheet.getName() !== expectedSheetName) {
    Logger.log(
      `Текущий лист не соответствует ожидаемому названию: ${expectedSheetName}`
    );
    return;
  }

  const range = sheet.getDataRange();
  const values = range.getValues();

  // Более разнообразные пастельные цвета
  const pastelColors = [
    "#FFB3BA", // Розовый
    "#FF677D", // Красный
    "#FFB74D", // Желтый
    "#A7D3A9", // Зеленый
    "#BAE1FF", // Голубой
    "#B39DDB", // Фиолетовый
    "#FFD1A0", // Персиковый
  ];

  // Создаем объект для хранения цветовой информации
  const colorMap = {};
  let colorIndex = 0;

  // Проходим по всем строкам, начиная со второй
  for (let i = 1; i < values.length; i++) {
    const dateValue = values[i][4]; // Столбец E (индекс 4)

    // Проверяем, является ли значение датой
    if (dateValue instanceof Date) {
      const dateString = dateValue.toDateString(); // Приводим дату к строке без учета времени

      // Если даты еще нет в colorMap, добавляем новую дату и цвет
      if (!colorMap[dateString]) {
        colorMap[dateString] = pastelColors[colorIndex % pastelColors.length];
        colorIndex++;
      }

      // Задаем цвет для строки
      const rowRange = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());
      rowRange.setBackground(colorMap[dateString]);
      // Устанавливаем границы для ячеек
      rowRange.setBorder(true, true, true, true, true, true);
    } else {
      const rowRange = sheet.getRange(i + 1, 1, 1, sheet.getLastColumn());
      rowRange.setBackground(null); // Убираем цвет, если значение не является датой
      rowRange.setBorder(true, true, true, true, true, true); // Устанавливаем границы
    }
  }
}
