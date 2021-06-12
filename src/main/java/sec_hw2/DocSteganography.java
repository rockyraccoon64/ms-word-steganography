package sec_hw2;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import java.io.*;
import java.nio.file.Files;
import java.util.List;

/* Информационная безопасность
 * Домашняя работа №2, вариант 2
 * Сало Андрей, МЕН-472201 (МО-401)
 * Синтаксис: java -jar Homework2.jar [txtFile.txt] [docxFile.docx] */

public class DocSteganography {

    public static void main(String[] args) {
        if (args.length != 2) {
            System.out.println("Error: wrong syntax");
            return;
        }
        File txtFile = new File(args[0]);
        File docFile = new File(args[1]);

        if (!txtFile.exists()) {
            System.out.println("Error: text file not found");
            return;
        }
        if (!docFile.exists()) {
            System.out.println("Error: doc file not found");
            return;
        }
        try {
            String text = getText(txtFile);
            XWPFDocument doc = getDocument(docFile);
            hideTextInDocFile(text, doc);
            writeDocument(doc, docFile);
        }
        catch (Exception ex) {
            System.out.println("Error: " + ex.getMessage());
            ex.printStackTrace();
        }
    }

    // Прочитать текст из txt-файла
    public static String getText(File file) throws IOException {
        return Files.readString(file.toPath());
    }

    // Получить структуру docx-файла
    public static XWPFDocument getDocument(File docFile) throws IOException {
        try (FileInputStream fis = new FileInputStream(docFile)) {
            return new XWPFDocument(fis);
        }
    }

    // Скрыть текст в docx-файле
    public static void hideTextInDocFile(String text, XWPFDocument doc)
            throws Exception {

        char[] hiddenTextChars = text.toLowerCase().toCharArray();
        int hiddenTextPosition = 0;

        List<XWPFParagraph> paragraphList = doc.getParagraphs();

        // Проходим по абзацам (Paragraph) документа
        OuterFor: for (XWPFParagraph p : paragraphList) {
            List<XWPFRun> runs = p.getRuns();

            // Проходим по последовательностям одинаково оформленных символов (Run) в абзаце (Paragraph)
            for (int runIndex = 0; runIndex < runs.size(); runIndex++) {
                XWPFRun run = runs.get(runIndex);
                String runText = run.getText(0);
                char[] runTextArray = runText.toLowerCase().toCharArray();

                boolean matchFoundInRun = false;
                int matchIndex = -1;
                int matchAmt = 0;

                // Ищем первую совпадающую последовательность символов в этом run
                for (int i = 0; i < runTextArray.length; i++) {

                    // Если уже весь текст вставлен, то выходим из текущего цикла
                    if (hiddenTextPosition >= hiddenTextChars.length)
                        break;

                    // Проверяем, совпадает ли текущий символ документа с текущим символом скрытого текста
                    boolean currentLetterMatches = (runTextArray[i] == hiddenTextChars[hiddenTextPosition]);
                    if (currentLetterMatches) {
                        if (!matchFoundInRun) {
                            // Если совпадает, то отмечаем его и продолжаем идти
                            matchFoundInRun = true;
                            matchIndex = i;
                            hiddenTextPosition++;
                            matchAmt++;
                        }
                        else {
                            // Если это не первый совпадающий символ подряд, то отмечаем их кол-во и продолжаем идти
                            matchAmt++;
                            hiddenTextPosition++;
                        }
                    }
                    else if (matchFoundInRun) {
                        // Если символы перестали совпадать, то выходим из цикла
                        break;
                    }
                }

                // Если не нашли совпадающих символов, переходим к следующему Run
                if (!matchFoundInRun)
                    continue;

                // Задаём новое оформление у последовательности совпадающих символов (увеличиваем размер текста на 1)
                int matchEndIndex = matchIndex + matchAmt;
                XWPFRun matchedRun = p.insertNewRun(runIndex + 1);
                copyRunProperties(run, matchedRun);
                matchedRun.setText(runText.substring(matchIndex, matchEndIndex));
                Double fontSize = run.getFontSizeAsDouble();
                if (fontSize == null) {
                    fontSize = doc.getStyles().getDefaultRunStyle().getFontSizeAsDouble();
                }
                matchedRun.setFontSize(fontSize + 1);

                // Если за совпадением следует ещё один фрагмент текста, оставляем у него оригинальное оформление
                if (matchEndIndex < runTextArray.length) {
                    XWPFRun followingRun = p.insertNewRun(runIndex + 2);
                    copyRunProperties(run, followingRun);
                    followingRun.setText(runText.substring(matchEndIndex, runTextArray.length));
                }

                // Если совпадение найдено в середине текста, то оставляем у первых символов оригинальное оформление
                if (matchIndex > 0) {
                    run.setText(runText.substring(0, matchIndex), 0);
                }
                else {
                    // Иначе просто удаляем пустой Run
                    p.removeRun(runIndex);
                    runIndex--;
                }
                runIndex++;

                // Если все символы текста найдены, заканчиваем обход документа
                if (hiddenTextPosition >= hiddenTextChars.length)
                    break OuterFor;
            }
        }

        // Если остались символы текста без совпадений, то текст нельзя скрыть в документе
        if (hiddenTextPosition < hiddenTextChars.length)
            throw new Exception("Can't hide text in docx: not enough symbols in docx file");
    }

    // Копируем свойства одной последовательности символов в другую
    public static void copyRunProperties(XWPFRun source, XWPFRun dest) {
        CTR ctrSource = source.getCTR();
        CTRPr rPrSource = ctrSource.getRPr();
        if (rPrSource != null) {
            CTR ctrDest = dest.getCTR();
            ctrDest.setRPr(rPrSource);
        }
    }

    // Записать изменения в docx-файл
    public static void writeDocument(XWPFDocument doc, File docFile) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(docFile, false)) {
            doc.write(fos);
        }
    }
}
