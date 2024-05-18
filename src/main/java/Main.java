import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) {
        // Step 1: Read file into list
        List<String> lines = readFile("data.txt");

        // Step 4: Create and populate Map<time, List<Program>>
        Map<BroadcastsTime, List<Program>> programMap = new TreeMap<>();
        String channel = null;
        BroadcastsTime time = null;
        for (String line : lines) {
            if (line.startsWith("#")) { // Assume channel name
                channel = line.substring(1).trim();
            } else {
                if (time == null) { // Expecting time
                    String[] timeParts = line.split(":");
                    byte hour = Byte.parseByte(timeParts[0]);
                    byte minutes = Byte.parseByte(timeParts[1]);
                    time = new BroadcastsTime(hour, minutes);
                } else { // Program name
                    Program program = new Program(channel, time, line.trim());
                    if (!programMap.containsKey(time)) {
                        programMap.put(time, new ArrayList<>());
                    }
                    programMap.get(time).add(program);
                    time = null; // Reset time for the next program
                }
            }
        }

        // Step 5: Create List<Program> with all programs
        List<Program> allPrograms = new ArrayList<>();
        for (List<Program> programs : programMap.values()) {
            allPrograms.addAll(programs);
        }

        // Step 6: Sort programs by channel, time
        Collections.sort(allPrograms, Comparator.comparing(Program::getChannel).thenComparing(Program::getTime));

        // Step 7: Print all programs
        System.out.println("All Programs:");
        for (Program program : allPrograms) {
            System.out.println(program.getChannel() + ", " + program.getTime().getHour() + ":" + program.getTime().getMinutes() + ", " + program.getName());
        }

        // Step 8: Find programs by name
        findProgramByName(allPrograms, "Слово пастыря");

        // Step 9: Find programs of a specific channel that are airing now
        findProgramsByChannelAiringNow(programMap, "Звезда");

        // Step 10: Find programs of a specific channel airing in a specific time range
        findProgramsByChannelInTimeRange(programMap, "Первый", new BroadcastsTime((byte) 10, (byte) 0), new BroadcastsTime((byte) 10, (byte) 15));

        // Step 11: Save sorted data to .xlsx file
        saveToExcel(allPrograms);
    }

    private static List<String> readFile(String fileName) {
        List<String> lines = new ArrayList<>();
        try (BufferedReader br = new BufferedReader(new FileReader(fileName))) {
            String line;
            while ((line = br.readLine()) != null) {
                lines.add(line);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return lines;
    }

    private static void findProgramByName(List<Program> programs, String name) {
        System.out.println("Program with name '" + name + "':");
        boolean found = false;
        for (Program program : programs) {
            if (program.getName().equalsIgnoreCase(name)) {
                found = true;
                System.out.println(program.getChannel() + ", " + program.getTime().getHour() + ":" + program.getTime().getMinutes() + ", " + program.getName());
            }
        }
        if (!found) {
            System.out.println("No programs found with name '" + name + "'.");
        }
    }

    private static void findProgramsByChannelAiringNow(Map<BroadcastsTime, List<Program>> programMap, String channel) {
        BroadcastsTime now = new BroadcastsTime((byte) 10, (byte) 30); // Предполагаем, что текущее время 10:30
        System.out.println("Programs airing now on channel '" + channel + "':");
        boolean found = false;

        BroadcastsTime prevEndTime = null;
        BroadcastsTime currentStartTime = null;
        for (Map.Entry<BroadcastsTime, List<Program>> entry : programMap.entrySet()) {
            BroadcastsTime startTime = entry.getKey();
            List<Program> programs = entry.getValue();

            for (Program program : programs) {
                if (program.getChannel().equalsIgnoreCase(channel)) {
                    // Если текущее время находится между временем окончания предыдущей программы и временем начала текущей программы
                    if (prevEndTime != null && prevEndTime.before(now) && startTime.after(now)) {
                        found = true;
                        System.out.println(program.getChannel() + ", " + currentStartTime.getHour() + ":" + currentStartTime.getMinutes() + ", " + program.getName());
                        return; // Останавливаем поиск после нахождения первого совпадения
                    }
                    prevEndTime = startTime; // Для упрощения используем время начала как время окончания предыдущей программы
                    currentStartTime = startTime;
                }
            }
        }

        if (!found) {
            System.out.println("No programs airing now on channel '" + channel + "'.");
        }
    }

    private static void findProgramsByChannelInTimeRange(Map<BroadcastsTime, List<Program>> programMap, String channel, BroadcastsTime start, BroadcastsTime end) {
        boolean found = false;
        // Уменьшаем время окончания на 1 минуту
        BroadcastsTime adjustedEndTime = new BroadcastsTime(end.getHour(), end.getMinutes());
        if (adjustedEndTime.getMinutes() == 0) {
            adjustedEndTime.setHour((byte) (adjustedEndTime.getHour() - 1));
            adjustedEndTime.setMinutes((byte) 59);
        } else {
            adjustedEndTime.setMinutes((byte) (adjustedEndTime.getMinutes() - 1));
        }

        System.out.println("Programs of channel '" + channel + "' airing between " + start.getHour() + ":" + start.getMinutes() + " and " + end.getHour() + ":" + end.getMinutes() + ":");

        for (Map.Entry<BroadcastsTime, List<Program>> entry : programMap.entrySet()) {
            BroadcastsTime time = entry.getKey();
            // Проверяем, попадает ли время программы в указанный интервал
            if (time.between(start, adjustedEndTime)) {
                for (Program program : entry.getValue()) {
                    if (program.getChannel().equalsIgnoreCase(channel)) {
                        found = true;
                        System.out.println(program.getChannel() + ", " + program.getTime().getHour() + ":" + program.getTime().getMinutes() + ", " + program.getName());
                    }
                }
            }
        }

        if (!found) {
            System.out.println("No programs found on channel '" + channel + "' in the specified time range.");
        }
    }

    private static void saveToExcel(List<Program> programs) {
        String[] columns = {"Channel", "Time", "Name"};

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Programs");

            // Header
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < columns.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns[i]);
            }

            // Data
            int rowNum = 1;
            for (Program program : programs) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(program.getChannel());
                row.createCell(1).setCellValue(program.getTime().getHour() + ":" + program.getTime().getMinutes());
                row.createCell(2).setCellValue(program.getName());
            }

            // Write to file
            try (FileOutputStream fileOut = new FileOutputStream("programs.xlsx")) {
                workbook.write(fileOut);
                System.out.println("Data saved to Excel file.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}