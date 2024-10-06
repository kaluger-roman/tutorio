/** @OnlyCurrentDoc */

const DAY_OF_WEEK = ["ПН", "ВТ", "СР", "ЧТ", "ПТ", "СБ", "ВС"];
const MONTHS = [
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

const LESSON_STATUS = {
  canceled: "Отменено",
  finished: "Завершено",
  moved: "Перенесено",
  planned: "Запланировано",
};

const FACT_LESSONS_COLUMN_INDEXES = {
  name: 0,
  grade: 1,
  goal: 2,
  price: 3,
  date: 4,
  duration: 5,
  isRegular: 6,
  status: 7,
  moveToDate: 8,
};

const CALENDAR_COLORS = {
  ege: CalendarApp.EventColor.RED,
  oge: CalendarApp.EventColor.ORANGE,
  school: CalendarApp.EventColor.BLUE,
  olimp: CalendarApp.EventColor.MAUVE,
};

const goalTagToColor = (tag) => {
  if (tag.includes("ЕГЭ")) return CALENDAR_COLORS.ege;
  if (tag.includes("ОГЭ")) return CALENDAR_COLORS.oge;
  if (tag.includes("Школа")) return CALENDAR_COLORS.school;
  if (tag.includes("Олимпиады")) return CALENDAR_COLORS.olimp;

  return CALENDAR_COLORS.school;
};

const convertDayToMs = (day = 0, hours = 0, minutes = 0, seconds = 0, ms = 0) =>
  day * 24 * 3600000 + hours * 3600000 + minutes * 60000 + seconds * 1000 + ms;

const timeToParts = (time) => {
  const [hh, mm, ss] =
    time
      ?.toString()
      ?.match(/\d{2}:\d{2}:\d{2}/g)?.[0]
      ?.split(":") || [];

  return { hh, mm, ss };
};

const buildEndDate = (startDate, duration) => {
  const durationMs = convertDayToMs(
    0,
    duration.getHours(),
    duration.getMinutes(),
    duration.getSeconds(),
    duration.getMilliseconds()
  );

  return new Date(startDate.getTime() + durationMs);
};

const buildEventTitle = ({ name, price, grade, goal, isMoved }) =>
  `${isMoved ? "(ПЕРЕНОС)" : ""}${name} ${grade ? `${grade} класс` : ""}  ${
    goal ? goal : ""
  } (${price})`;

const getSheetMainTableValues = (sheetName) => {
  const [headers, ...rows] = spreadsheet
    .getSheetByName(sheetName)
    .getRange(1, 1)
    .getDataRegion()
    .getValues();

  return [headers, rows];
};

const getFactLessonsTableValues = (monthNum) => {
  const [headers, rows] = getSheetMainTableValues(
    buildFactLessonsSheetName(monthNum)
  );

  return {
    headers: factLessonRowToData(headers),
    rows: rows.map((x, index) => ({
      ...factLessonRowToData(x),
      rowIndex: index + 2,
    })), // +2 because of headers and 1-based index
  };
};

const convertDayOfWeekToRuOrder = (date) => {
  const dayOfWeek = new Date().getDay();

  return dayOfWeek === 0 ? 6 : dayOfWeek - 1;
};

const buildCurrentAndNextMonth = () => {
  const currentMonth = new Date().getMonth();
  const nextMonth = currentMonth + 1;

  return [currentMonth, nextMonth];
};

const getNextWeekEnd = () => {
  const nextWeekEnd = new Date();

  nextWeekEnd.setDate(
    new Date().getDate() - convertDayOfWeekToRuOrder(new Date()) + 13
  ); // 13 is to get next sunday from curr monday
  nextWeekEnd.setHours(23, 59, 59, 999);

  return nextWeekEnd;
};

const getCurrentMonthStart = () =>
  new Date(new Date().getFullYear(), new Date().getMonth(), 1);

const protectRange = (range) => {
  const protection = range.protect();

  protection.removeEditors(protection.getEditors());
  protection.setWarningOnly(true);
};

const scheduleRowToData = (row) => ({
  name: row[0],
  grade: row[1],
  goal: row[2],
  price: row[3],
  dayOfWeek: DAY_OF_WEEK.findIndex((x) => x === row[4]),
  startTime: row[5],
  duration: row[6],
});

const factLessonRowToData = (row) => ({
  name: row[FACT_LESSONS_COLUMN_INDEXES.name],
  grade: row[FACT_LESSONS_COLUMN_INDEXES.grade],
  goal: row[FACT_LESSONS_COLUMN_INDEXES.goal],
  price: row[FACT_LESSONS_COLUMN_INDEXES.price],
  date: row[FACT_LESSONS_COLUMN_INDEXES.date],
  duration: row[FACT_LESSONS_COLUMN_INDEXES.duration],
  isRegular: row[FACT_LESSONS_COLUMN_INDEXES.isRegular],
  status: row[FACT_LESSONS_COLUMN_INDEXES.status],
  moveToDate: row[FACT_LESSONS_COLUMN_INDEXES.moveToDate],
});

const buildFactLessonsSheetName = (monthNum) => `Факт. ${MONTHS[monthNum]}`;

const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const calendar = CalendarApp.getDefaultCalendar();

const getLessonNearestDates = (row) => {
  const dayOfWeekRu = convertDayOfWeekToRuOrder(new Date());

  const { dayOfWeek, startTime, duration } = scheduleRowToData(row);

  if (!isFinite(dayOfWeek) || !startTime || !duration) return [];

  const diffNearestLesson =
    dayOfWeek < dayOfWeekRu
      ? DAY_OF_WEEK.length - dayOfWeekRu + dayOfWeek
      : dayOfWeek - dayOfWeekRu;

  const { hh, mm, ss } = timeToParts(startTime);

  const date = new Date(
    new Date().getTime() + convertDayToMs(diffNearestLesson)
  );

  date.setHours(hh, mm, ss, 0);

  const nextDate = new Date(date.getTime() + convertDayToMs(7));

  if (getNextWeekEnd().getTime() < nextDate.getTime()) return [date];

  return [date, nextDate];
};

const createNextMonthFactLessonsTemplate = () => {
  const [currentMonth, nextMonth] = buildCurrentAndNextMonth();

  const templateSheet = spreadsheet.getSheetByName("Факт. Шаблон.");

  const createTemplate = (monthNum) => {
    const factLessonsSheetName = buildFactLessonsSheetName(monthNum);

    if (!spreadsheet.getSheetByName(factLessonsSheetName)) {
      const sheet = templateSheet.copyTo(spreadsheet);

      sheet.setName(factLessonsSheetName);

      const { headers: headerRangeValues } =
        getFactLessonsTableValues(monthNum);

      protectRange(
        sheet.getRange(1, 1, 1, Object.values(headerRangeValues).length)
      );
    }
  };

  createTemplate(currentMonth);
  createTemplate(nextMonth);
};

const createFactLessonsFromSchedule = () => {
  const [, scheduleRows] = getSheetMainTableValues("Расписание");

  for (const scheduleRow of scheduleRows) {
    const nearestDates = getLessonNearestDates(scheduleRow);
    const { name, price, duration, goal, grade } =
      scheduleRowToData(scheduleRow);

    for (const nearestDate of nearestDates) {
      const dateSheetName = buildFactLessonsSheetName(nearestDate.getMonth());
      const dateSheet = spreadsheet.getSheetByName(dateSheetName);
      const { headers, rows: factLessonsValues } = getFactLessonsTableValues(
        nearestDate.getMonth()
      );

      const isExist = factLessonsValues.find(
        ({ date }) => date.getTime?.() === nearestDate.getTime()
      );

      if (!isExist && duration && name && price) {
        const newRowIndex = factLessonsValues.length + 2; // 1 is header row and 1 is 1 based system

        dateSheet.insertRowBefore(newRowIndex);
        const newRowRange = dateSheet.getRange(
          newRowIndex,
          1,
          1,
          Object.values(headers).length
        );

        newRowRange.setValues([
          [
            name,
            grade,
            goal,
            price,
            nearestDate,
            duration,
            true,
            LESSON_STATUS.planned,
            "",
          ],
        ]);

        protectRange(newRowRange);
      }
    }
  }
};

const markEndnedLessons = () => {
  const [currentMonth] = buildCurrentAndNextMonth();
  const { rows: currentMonthFactLessons } =
    getFactLessonsTableValues(currentMonth);

  const newFinishedLessons = currentMonthFactLessons.filter(
    ({ date, status, duration, moveToDate }) => {
      const lessonDate = status === LESSON_STATUS.moved ? moveToDate : date;

      return (
        buildEndDate(lessonDate, duration)?.getTime() < new Date().getTime() &&
        ![LESSON_STATUS.finished, LESSON_STATUS.canceled].includes(status)
      );
    }
  );

  newFinishedLessons.forEach(({ rowIndex }) => {
    const sheet = spreadsheet.getSheetByName(
      buildFactLessonsSheetName(currentMonth)
    );

    sheet
      .getRange(rowIndex, FACT_LESSONS_COLUMN_INDEXES.status + 1)
      .setValue(LESSON_STATUS.finished);
  });
};

const updateGoogleCalendarEvents = () => {
  const nearEvents = calendar.getEvents(
    getCurrentMonthStart(),
    getNextWeekEnd()
  );

  const [currentMonth, nextMonth] = buildCurrentAndNextMonth();
  const { rows: currentMonthFactLessons } =
    getFactLessonsTableValues(currentMonth);
  const { rows: nextMonthFactLessons } = getFactLessonsTableValues(nextMonth);
  const nearestLessons = [...currentMonthFactLessons, ...nextMonthFactLessons];

  for (const lesson of nearestLessons) {
    const { name, price, date, duration, grade, goal } = lesson;
    const eventTitle = buildEventTitle({ name, price, grade, goal });
    const isExist = nearEvents.find(
      (x) =>
        x.getStartTime().getTime() === date?.getTime?.() &&
        x.getTitle()?.includes(eventTitle)
    );

    if (!isExist && name && price && date && duration) {
      const endDate = buildEndDate(date, duration);

      const event = calendar.createEvent(eventTitle, date, endDate);

      if (goal) event.setColor(goalTagToColor(goal));
    }
  }

  for (const event of nearEvents) {
    const eventTitle = event.getTitle();
    const isExist = nearestLessons.find(
      ({ name, price, date, duration, grade, goal }) =>
        event.getStartTime().getTime() === date?.getTime?.() &&
        eventTitle?.includes(buildEventTitle({ name, price, grade, goal }))
    );

    if (!isExist) event.deleteEvent();
  }
};

const deleteCanceledEvents = () => {
  const [currentMonth, nextMonth] = buildCurrentAndNextMonth();
  const { rows: currentMonthFactLessons } =
    getFactLessonsTableValues(currentMonth);
  const { rows: nextMonthFactLessons } = getFactLessonsTableValues(nextMonth);

  const canceledLessons = [
    ...currentMonthFactLessons,
    ...nextMonthFactLessons,
  ].filter(({ status }) => status === LESSON_STATUS.canceled);

  const nearEvents = calendar.getEvents(
    getCurrentMonthStart(),
    getNextWeekEnd()
  );

  for (const lesson of canceledLessons) {
    const { name, price, date, grade, goal, moveToDate } = lesson;
    const eventTitleActive = buildEventTitle({ name, price, grade, goal });
    const event = nearEvents.find(
      (x) =>
        x.getStartTime().getTime() === (date || moveToDate)?.getTime?.() &&
        x.getTitle()?.includes(eventTitleActive)
    );

    event.deleteEvent();
  }
};

const syncSchedule = () => {
  createNextMonthFactLessonsTemplate();
  createFactLessonsFromSchedule();
  markEndnedLessons();
  updateGoogleCalendarEvents();
  deleteCanceledEvents();
};
