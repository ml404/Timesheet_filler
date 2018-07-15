package com.layton.matthew;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.kohsuke.args4j.CmdLineException;
import org.kohsuke.args4j.CmdLineParser;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalField;
import java.time.temporal.WeekFields;
import java.util.Locale;

public class TimesheetWriterMain {
    public static void main(String[] args) {
        createDocument(args);
    }

    public static void createDocument(String[] args) {
        TimesheetReaderOptions timesheetReaderOptions = new TimesheetReaderOptions();
        CmdLineParser parser = new CmdLineParser(timesheetReaderOptions);
        try {
            parser.parseArgument(args);
            TimesheetReader timesheetReader = new TimesheetReader();
            XWPFDocument xwpfDocument = timesheetReader.readTimeSheet(timesheetReaderOptions.readLocation);
            setFirstTable(xwpfDocument);
            setSecondTable(timesheetReaderOptions, xwpfDocument);
            xwpfDocument.write(new FileOutputStream(new File(timesheetReaderOptions.writeLocation)));
        } catch (CmdLineException e) {
            // handling of wrong arguments
            System.err.println(e.getMessage());
            parser.printUsage(System.err);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void setFirstTable(XWPFDocument doc) {
        XWPFTable tableArray = doc.getTableArray(0);

        //Set my name, period covered, client site and current week
        XWPFTableRow row1 = tableArray.getRow(0);
        row1.getCell(1).removeParagraph(0);
        row1.getCell(1).setText("Matthew Grant Layton");

        String periodCovered = getPeriodCovered();
        row1.getCell(3).removeParagraph(0);
        row1.getCell(3).setText(periodCovered);

        XWPFTableRow row2 = tableArray.getRow(1);
        row2.getCell(1).removeParagraph(0);
        row2.getCell(1).setText("Standard Chartered Bank");

        row2.getCell(3).removeParagraph(0);
        row2.getCell(3).setText(String.valueOf(getCurrentWeek()));

    }

    private static void setSecondTable(TimesheetReaderOptions timesheetReaderOptions, XWPFDocument doc) {
        int weeklyWorkedHours = 0;
        int weeklyWorkedMinutes = 0;


        XWPFTable tableArray = doc.getTableArray(1);

        //Set start time, lunch time taken and time finished
        for (int i = 1; i <= 5; i++) {
            XWPFTableRow startTimesRow = tableArray.getRow(1);
            startTimesRow.getCell(i).removeParagraph(0);
            startTimesRow.getCell(i).setText(timesheetReaderOptions.startTimes.get(i - 1));

            XWPFTableRow endTimesRow = tableArray.getRow(2);
            endTimesRow.getCell(i).removeParagraph(0);
            endTimesRow.getCell(i).setText(timesheetReaderOptions.endTimes.get(i - 1));

            XWPFTableRow lunchTimesRow = tableArray.getRow(3);
            lunchTimesRow.getCell(i).removeParagraph(0);
            lunchTimesRow.getCell(i).setText(timesheetReaderOptions.lunchTimes.get(i - 1));

            int startHour = Integer.parseInt(timesheetReaderOptions.startTimes.get(i - 1).split(":")[0]);
            int startMinute = Integer.parseInt(timesheetReaderOptions.startTimes.get(i - 1).split(":")[1]);

            int endHour = Integer.parseInt(timesheetReaderOptions.endTimes.get(i - 1).split(":")[0]);
            int endMinute = Integer.parseInt(timesheetReaderOptions.endTimes.get(i - 1).split(":")[1]);

            int lunchHour = Integer.parseInt(timesheetReaderOptions.lunchTimes.get(i - 1).split(":")[0]);
            int lunchMinute = Integer.parseInt(timesheetReaderOptions.lunchTimes.get(i - 1).split(":")[1]);

            int dailyWorkedHours = (endHour - startHour - lunchHour);
            int dailyWorkedMinutes = (endMinute - startMinute - lunchMinute);

            while (dailyWorkedMinutes < 0) {
                dailyWorkedHours--;
                dailyWorkedMinutes += 60;
            }
            while (dailyWorkedMinutes > 60) {
                dailyWorkedHours++;
                dailyWorkedMinutes -= 60;
            }

            weeklyWorkedHours += dailyWorkedHours;
            weeklyWorkedMinutes += dailyWorkedMinutes;
            XWPFTableRow workedTimes = tableArray.getRow(4);
            workedTimes.getCell(i).removeParagraph(0);
            workedTimes.getCell(i).setText(String.format("%sh %sm", dailyWorkedHours, dailyWorkedMinutes));
        }

        while (weeklyWorkedMinutes < 0) {
            weeklyWorkedHours--;
            weeklyWorkedMinutes += 60;
        }
        while (weeklyWorkedMinutes > 60) {
            weeklyWorkedHours++;
            weeklyWorkedMinutes -= 60;
        }

        XWPFTableRow workedTimes = tableArray.getRow(4);
        workedTimes.getCell(8).removeParagraph(0);
        workedTimes.getCell(8).setText(String.format("%sh %sm", weeklyWorkedHours, weeklyWorkedMinutes));
    }

    private static String getPeriodCovered() {
        LocalDate today = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/YYYY");
        // Go backward to get Monday
        LocalDate monday = today;
        while (monday.getDayOfWeek() != DayOfWeek.MONDAY) {
            monday = monday.minusDays(1);
        }
        // Go forward to get Friday
        LocalDate friday = today;
        while (friday.getDayOfWeek() != DayOfWeek.SUNDAY) {
            friday = friday.plusDays(1);
        }
        return monday.format(formatter) + " - " + friday.format(formatter);
    }

    private static int getCurrentWeek() {
        LocalDate date = LocalDate.now();
        TemporalField woy = WeekFields.of(Locale.getDefault()).weekOfWeekBasedYear();
        return date.get(woy);
    }
}
