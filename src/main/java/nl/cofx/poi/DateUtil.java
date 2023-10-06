package nl.cofx.poi;

import lombok.experimental.UtilityClass;

import java.time.DayOfWeek;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAdjusters;
import java.time.temporal.TemporalField;
import java.time.temporal.WeekFields;


@UtilityClass
public class DateUtil {

    private static final int SEVEN_DAYS = 7;
    private static final WeekFields WEEK_FIELDS = WeekFields.of(DayOfWeek.MONDAY, SEVEN_DAYS);
    private static final TemporalField DAY_OF_WEEK = WEEK_FIELDS.dayOfWeek();
    private static final TemporalField WEEK_OF_YEAR = WEEK_FIELDS.weekOfWeekBasedYear();

    public Instant midnight(LocalDate date, ZoneId zoneId) {
        return date.atStartOfDay(zoneId).toInstant();
    }

    public String monthName(LocalDate date) {
        var thursdayInSameWeek = date;

        if (DayOfWeek.THURSDAY.compareTo(date.getDayOfWeek()) < 0) {
            thursdayInSameWeek = date.with(TemporalAdjusters.previous(DayOfWeek.THURSDAY));
        } else if (DayOfWeek.THURSDAY.compareTo(date.getDayOfWeek()) > 0) {
            thursdayInSameWeek = date.with(TemporalAdjusters.next(DayOfWeek.THURSDAY));
        }

        return thursdayInSameWeek.format(DateTimeFormatter.ofPattern("MMMM"));
    }

    public int weekNumber(LocalDate date) {
        return date.get(WEEK_OF_YEAR);
    }

    public LocalDate firstDayOfWeek(LocalDate date) {
        return date.with(DAY_OF_WEEK, 1);
    }

    public LocalDate firstDayOfWeek(Instant instant, ZoneId zoneId) {
        return instant.atZone(zoneId).with(DAY_OF_WEEK, 1).toLocalDate();
    }

    public LocalDate firstDayOfNextWeek(Instant instant, ZoneId zoneId) {
        return firstDayOfWeek(instant, zoneId).plusWeeks(1);
    }
}
