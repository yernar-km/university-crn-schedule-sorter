
import java.util.List;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;





public class CRNFinder {
    private static String filePath = null; // Глобальный путь к файлу

    public static void main(String[] args) {
        SwingUtilities.invokeLater(CRNFinder::createAndShowGUI);
    }

    private static void createAndShowGUI() {
        JFrame frame = new JFrame("Поиск по CRN");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(600, 450);
        frame.setLayout(new BorderLayout());

        JPanel panel = new JPanel(new GridLayout(3, 1));
        JTextField crnInputField = new JTextField();
        crnInputField.setToolTipText("Введите CRN (через запятую)");

        panel.add(new JLabel("Введите CRN (через запятую):"));
        panel.add(crnInputField);

        JPanel filePanel = new JPanel(new FlowLayout());
        JButton chooseFileButton = new JButton("Выбрать файл");
        JLabel fileLabel = new JLabel("Файл не выбран");

        filePanel.add(chooseFileButton);
        filePanel.add(fileLabel);
        panel.add(filePanel);

        JButton searchButton = new JButton("Найти");

        JTextArea resultArea = new JTextArea();
        resultArea.setEditable(false);
        resultArea.setLineWrap(true);
        JScrollPane scrollPane = new JScrollPane(resultArea);

        frame.add(panel, BorderLayout.NORTH);
        frame.add(scrollPane, BorderLayout.CENTER);
        frame.add(searchButton, BorderLayout.SOUTH);

        // Обработчик выбора файла
        chooseFileButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();
            int result = fileChooser.showOpenDialog(null);
            if (result == JFileChooser.APPROVE_OPTION) {
                filePath = fileChooser.getSelectedFile().getAbsolutePath();
                fileLabel.setText("Выбран: " + fileChooser.getSelectedFile().getName());
            }
        });

        // Обработчик поиска
        searchButton.addActionListener(e -> {
            if (filePath == null) {
                resultArea.setText("Сначала выберите файл.");
                return;
            }

            String crnInput = crnInputField.getText().trim();
            if (crnInput.isEmpty()) {
                resultArea.setText("Введите хотя бы один CRN.");
                return;
            }

            List<String> crnList = Arrays.asList(crnInput.split(","));
            resultArea.setText("");

            try (FileInputStream fis = new FileInputStream(filePath); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0);
                Map<String, List<Row>> results = new HashMap<>();

                for (Row row : sheet) {
                    Cell crnCell = row.getCell(5); // CRN в 6-м столбце
                    if (crnCell != null) {
                        String cellValue = crnCell.toString().trim();
                        if (crnList.contains(cellValue)) {
                            results.putIfAbsent(cellValue, new ArrayList<>());
                            results.get(cellValue).add(row);
                        }
                    }
                }

                // Сортировка CRN по самой ранней дате
                crnList.sort((crn1, crn2) -> {
                    List<Row> rows1 = results.get(crn1);
                    List<Row> rows2 = results.get(crn2);

                    String earliestDate1 = rows1 == null ? "Нет данных" : rows1.stream()
                            .map(row -> getCellValue(row, 11))
                            .min(CRNFinder::compareDates)
                            .orElse("Нет данных");

                    String earliestDate2 = rows2 == null ? "Нет данных" : rows2.stream()
                            .map(row -> getCellValue(row, 11))
                            .min(CRNFinder::compareDates)
                            .orElse("Нет данных");

                    return compareDates(earliestDate1, earliestDate2);
                });

                for (String crn : crnList) {
                    resultArea.append("----------------------------------------------------------------------------" + "\n");
                    if (results.containsKey(crn)) {
                        List<Row> rows = results.get(crn);
                        rows.sort((row1, row2) -> compareDates(getCellValue(row1, 11), getCellValue(row2, 11)));
                        for (Row row : rows) {
                            resultArea.append(formatRowData(row) + "\n");
                        }
                    } else {
                        resultArea.append("Нет данных для CRN: " + crn + "\n");
                    }
                    resultArea.append("\n");
                }

            } catch (IOException ex) {
                resultArea.setText("Ошибка чтения файла: " + ex.getMessage());
            }
        });

        frame.setVisible(true);
    }

    private static String formatRowData(Row row) {
        return String.format("Teacher: %s\nDate: %s\nTime: %s\nRoom: %s",
                getCellValue(row, 1),
                getCellValue(row, 11),
                getCellValue(row, 13),
                getCellValue(row, 14));
    }

    private static String getCellValue(Row row, int cellIndex) {
        Cell cell = row.getCell(cellIndex);
        if (cell != null) {
            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                Date date = cell.getDateCellValue();
                SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy", Locale.forLanguageTag("ru"));
                return sdf.format(date);
            } else {
                return cell.toString().trim();
            }
        }
        return "Нет данных";
    }

    private static int compareDates(String date1, String date2) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy", Locale.ENGLISH);
        try {
            Date d1 = dateFormat.parse(date1);
            Date d2 = dateFormat.parse(date2);
            return d1.compareTo(d2);
        } catch (ParseException e) {
            return 0;
        }
    }
}
