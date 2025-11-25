package Service;

import org.apache.poi.xwpf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalAlignRun;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.stream.Collectors;

public class SpellCheckProcessor {

    private static final String API_URL = "http://localhost:11434/v1/chat/completions";
    private static final Set<String> STYLES_TO_SKIP = Set.of("CL", "AU", "EH", "TY", "DOI", "LRH", "RRH", "AF", "AT",
            "AS", "ABKWH", "ABKW", "H1", "cit", "AQ", "H2", "AN", "author", "adate", "atl", "stl", "vol", "iss", "first-page",
            "last-page", "REF", "org", "btl", "city", "pub", "aulabel", "Hyperlink", "CP", "H3", "DR", "Front matter",
            "OQ", "QS", "H4", "H5", "EX", "DI", "PO", "EQ", "EN", "NNUM", "CPB", "TCH", "TT", "TNL", "TBL", "CPSO");
    private static final String INPUT_FOLDER = "D:/before";
    private static final String OUTPUT_FOLDER = "D:/after";

    // Helper class to hold masked text and placeholders
    private static class TextWithPlaceholders {
        final String maskedText;
        final List<PlaceholderInfo> placeholders;

        TextWithPlaceholders(String maskedText, List<PlaceholderInfo> placeholders) {
            this.maskedText = maskedText;
            this.placeholders = new ArrayList<>(placeholders);
        }
    }

    // Store both the original text and its type (superscript/subscript)
    private static class PlaceholderInfo {
        final String originalText;
        final boolean isSuperscript;
        final boolean isSubscript;
        final RunFormatting formatting;

        PlaceholderInfo(String originalText, boolean isSuperscript, boolean isSubscript, RunFormatting formatting) {
            this.originalText = originalText;
            this.isSuperscript = isSuperscript;
            this.isSubscript = isSubscript;
            this.formatting = formatting;
        }
    }

    // Helper class to store run information
    private static class RunInfo {
        final String text;
        final boolean isSuperscript;
        final boolean isSubscript;
        final RunFormatting formatting;

        RunInfo(String text, boolean isSuperscript, boolean isSubscript, RunFormatting formatting) {
            this.text = text;
            this.isSuperscript = isSuperscript;
            this.isSubscript = isSubscript;
            this.formatting = formatting;
        }
    }

    // Store formatting properties separately
    private static class RunFormatting {
        CTRPr rPr; // Store the complete run properties

        static RunFormatting from(XWPFRun run) {
            RunFormatting fmt = new RunFormatting();
            try {
                CTR ctr = run.getCTR();
                if (ctr != null && ctr.isSetRPr()) {
                    fmt.rPr = (CTRPr) ctr.getRPr().copy();
                }
            } catch (Exception e) {
                fmt.rPr = null;
            }
            return fmt;
        }

        void applyTo(XWPFRun run, boolean forceGreen, boolean forceRedStrikethrough) {
            // First, copy all original formatting properties from CTRPr if available
            if (rPr != null) {
                try {
                    CTR ctr = run.getCTR();
                    if (ctr != null) {
                        CTRPr newRPr = (CTRPr) rPr.copy();
                        ctr.setRPr(newRPr);
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            
            // Then apply color changes for track changes (AFTER copying formatting)
            try {
                if (forceRedStrikethrough) {
                    run.setStrikeThrough(true);
                    run.setColor("2F4F4F"); // Red for deletions
                } else if (forceGreen) {
                    run.setColor("2CFF05"); // 652A0E Green for additions
                }
            } catch (Exception ignored) {}
        }
    }

    // Helper class to represent text differences
    private static class TextSegment {
        final String text;
        final boolean isChanged;
        final boolean isDeleted;
        final boolean isSuperscript;
        final boolean isSubscript;
        final RunFormatting formatting;

        TextSegment(String text, boolean isChanged, boolean isDeleted, boolean isSuperscript, boolean isSubscript, RunFormatting formatting) {
            this.text = text;
            this.isChanged = isChanged;
            this.isDeleted = isDeleted;
            this.isSuperscript = isSuperscript;
            this.isSubscript = isSubscript;
            this.formatting = formatting;
        }
    }

    // Diff operation types
    private enum DiffType {
        UNCHANGED, ADDED, REMOVED
    }

    private static class DiffResult {
        final String text;
        final DiffType type;

        DiffResult(String text, DiffType type) {
            this.text = text;
            this.type = type;
        }
    }

    public static void main(String[] args) {
        processFolder();
    }

    public static void processFolder() {
        try {
            Path inputPath = Paths.get(INPUT_FOLDER);
            Path outputPath = Paths.get(OUTPUT_FOLDER);

            if (!Files.exists(inputPath)) {
                Files.createDirectories(inputPath);
                System.out.println("Created input directory: " + INPUT_FOLDER);
            }

            if (!Files.exists(outputPath)) {
                Files.createDirectories(outputPath);
                System.out.println("Created output directory: " + OUTPUT_FOLDER);
            }

            List<Path> docxFiles = Files.list(inputPath)
                    .filter(path -> path.toString().toLowerCase().endsWith(".docx"))
                    .collect(Collectors.toList());

            if (docxFiles.isEmpty()) {
                System.out.println("No DOCX files found in " + INPUT_FOLDER);
                return;
            }

            for (Path docxFile : docxFiles) {
                try (InputStream inputStream = Files.newInputStream(docxFile)) {
                    String correctedFileName = readAndProcessDocxFile(inputStream, docxFile.getFileName().toString());
                    System.out.println("Processed: " + docxFile.getFileName() + " → " + correctedFileName);

                    // Move original to output folder
                    Path destinationPath = Paths.get(OUTPUT_FOLDER, docxFile.getFileName().toString());
                    Files.move(docxFile, destinationPath, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                    System.out.println("Moved original file to: " + destinationPath);
                } catch (Exception e) {
                    System.err.println("Error processing " + docxFile.getFileName() + ": " + e.getMessage());
                    e.printStackTrace();
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static String readAndProcessDocxFile(InputStream inputStream, String originalFileName) throws Exception {
        XWPFDocument doc = new XWPFDocument(inputStream);

        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            if (isParagraphStyleToSkip(paragraph)) {
                continue;
            }

            // STEP 1: Add markers for all superscripts/subscripts FIRST
            addMarkersToSuperSubscripts(paragraph);
            
            // STEP 2: Now proceed with grammar checking on the marked text
            List<RunInfo> runInfos = extractRunInfos(paragraph);
            if (runInfos.isEmpty()) {
                continue;
            }

            StringBuilder fullText = new StringBuilder();
            for (RunInfo info : runInfos) {
                fullText.append(info.text);
            }

            String originalText = fullText.toString();
            if (originalText.trim().isEmpty()) {
                continue;
            }

            TextWithPlaceholders masked = maskFromRunInfos(runInfos);
            String correctedMasked = callGrammarCheckApi(masked.maskedText);
            String correctedText = restorePlaceholders(correctedMasked, masked.placeholders);

            if (!originalText.equals(correctedText)) {
                rebuildParagraphWithChanges(paragraph, originalText, correctedText, runInfos, masked.placeholders);
            }
        }

        return writeToFile(doc, originalFileName);
    }

    /**
     * Add (SUP) or (SUB) markers after all superscripts and subscripts in a paragraph
     */
    private static void addMarkersToSuperSubscripts(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        List<Integer> insertPositions = new ArrayList<>();
        List<String> markerTexts = new ArrayList<>();
        
        // First pass: identify where to insert markers
        for (int i = 0; i < runs.size(); i++) {
            XWPFRun run = runs.get(i);
            String text = run.getText(0);
            if (text == null || text.isEmpty()) continue;
            
            boolean isSuperscript = isSuperscriptRun(run);
            boolean isSubscript = isSubscriptRun(run);
            
            // Check for unicode superscripts/subscripts in the text
            boolean hasUnicodeSuperSub = false;
            for (char c : text.toCharArray()) {
                if (isSuperscriptOrSubscript(c)) {
                    hasUnicodeSuperSub = true;
                    break;
                }
            }
            
            if (isSuperscript || isSubscript || hasUnicodeSuperSub) {
                insertPositions.add(i + 1); // Insert after this run
                
                if (isSuperscript) {
                    markerTexts.add("(SUP)");
                } else if (isSubscript) {
                    markerTexts.add("(SUB)");
                } else {
                    // For unicode, determine type from first character
                    boolean isUniSuper = isSuperscriptChar(text.charAt(0));
                    markerTexts.add(isUniSuper ? "(SUP)" : "(SUB)");
                }
            }
        }
        
        // Second pass: insert markers (in reverse order to maintain correct positions)
        for (int i = insertPositions.size() - 1; i >= 0; i--) {
            int position = insertPositions.get(i);
            String markerText = markerTexts.get(i);
            
            // Adjust position for already inserted markers
            for (int j = i + 1; j < insertPositions.size(); j++) {
                if (insertPositions.get(j) <= position) {
                    position++;
                }
            }
            
            XWPFRun markerRun = paragraph.insertNewRun(position);
            markerRun.setText(markerText, 0);
            markerRun.setColor("FF6600"); // Orange color for visibility
            markerRun.setBold(true);
            markerRun.setFontSize(10);
        }
    }

    private static List<RunInfo> extractRunInfos(XWPFParagraph paragraph) {
        List<RunInfo> runInfos = new ArrayList<>();
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.getText(0);
            if (text == null) text = "";
            boolean isSuperscript = isSuperscriptRun(run);
            boolean isSubscript = isSubscriptRun(run);
            RunFormatting formatting = RunFormatting.from(run);
            runInfos.add(new RunInfo(text, isSuperscript, isSubscript, formatting));
        }
        return runInfos;
    }

    private static boolean isSuperscriptRun(XWPFRun run) {
        try {
            CTR ctr = run.getCTR();
            if (ctr != null && ctr.isSetRPr() && ctr.getRPr().isSetVertAlign()) {
                STVerticalAlignRun.Enum vertAlign = ctr.getRPr().getVertAlign().getVal();
                return vertAlign == STVerticalAlignRun.SUPERSCRIPT;
            }
        } catch (Exception e) {
            // Ignore
        }
        return false;
    }

    private static boolean isSubscriptRun(XWPFRun run) {
        try {
            CTR ctr = run.getCTR();
            if (ctr != null && ctr.isSetRPr() && ctr.getRPr().isSetVertAlign()) {
                STVerticalAlignRun.Enum vertAlign = ctr.getRPr().getVertAlign().getVal();
                return vertAlign == STVerticalAlignRun.SUBSCRIPT;
            }
        } catch (Exception e) {
            // Ignore
        }
        return false;
    }

    private static TextWithPlaceholders maskFromRunInfos(List<RunInfo> runInfos) {
        List<PlaceholderInfo> placeholders = new ArrayList<>();
        StringBuilder masked = new StringBuilder();

        for (RunInfo info : runInfos) {
            if (info.isSuperscript || info.isSubscript) {
                placeholders.add(new PlaceholderInfo(info.text, info.isSuperscript, info.isSubscript, info.formatting));
                masked.append("«SUPSUB_").append(placeholders.size() - 1).append("»");
            } else {
                masked.append(maskUnicodeSuperSubscripts(info.text, placeholders, info.formatting));
            }
        }

        return new TextWithPlaceholders(masked.toString(), placeholders);
    }

    private static String maskUnicodeSuperSubscripts(String text, List<PlaceholderInfo> placeholders, RunFormatting formatting) {
        StringBuilder result = new StringBuilder();
        StringBuilder currentGroup = new StringBuilder();

        for (char c : text.toCharArray()) {
            if (isSuperscriptOrSubscript(c)) {
                currentGroup.append(c);
            } else {
                if (currentGroup.length() > 0) {
                    boolean isSuper = isSuperscriptChar(currentGroup.charAt(0));
                    placeholders.add(new PlaceholderInfo(currentGroup.toString(), isSuper, !isSuper, formatting));
                    result.append("«SUPSUB_").append(placeholders.size() - 1).append("»");
                    currentGroup.setLength(0);
                }
                result.append(c);
            }
        }

        if (currentGroup.length() > 0) {
            boolean isSuper = isSuperscriptChar(currentGroup.charAt(0));
            placeholders.add(new PlaceholderInfo(currentGroup.toString(), isSuper, !isSuper, formatting));
            result.append("«SUPSUB_").append(placeholders.size() - 1).append("»");
        }

        return result.toString();
    }

    private static boolean isSuperscriptOrSubscript(char c) {
        return (c >= 0x2070 && c <= 0x207F) || (c >= 0x2080 && c <= 0x208F);
    }

    private static boolean isSuperscriptChar(char c) {
        return c >= 0x2070 && c <= 0x207F;
    }

    private static String restorePlaceholders(String maskedText, List<PlaceholderInfo> placeholders) {
        StringBuilder result = new StringBuilder();
        int i = 0;

        while (i < maskedText.length()) {
            if (maskedText.startsWith("«SUPSUB_", i)) {
                int end = maskedText.indexOf("»", i);
                if (end != -1) {
                    String numPart = maskedText.substring(i + 8, end);
                    try {
                        int index = Integer.parseInt(numPart);
                        if (index >= 0 && index < placeholders.size()) {
                            result.append(placeholders.get(index).originalText);
                        } else {
                            result.append("«SUPSUB_").append(numPart).append("»");
                        }
                    } catch (NumberFormatException e) {
                        result.append("«SUPSUB_").append(numPart).append("»");
                    }
                    i = end + 1;
                } else {
                    result.append(maskedText.charAt(i));
                    i++;
                }
            } else {
                result.append(maskedText.charAt(i));
                i++;
            }
        }

        return result.toString();
    }

    private static void rebuildParagraphWithChanges(XWPFParagraph paragraph, String originalText,
                                                     String correctedText, List<RunInfo> originalRuns,
                                                     List<PlaceholderInfo> placeholders) {
        RunFormatting defaultFormatting = originalRuns.isEmpty() ? new RunFormatting() : originalRuns.get(0).formatting;

        for (int i = paragraph.getRuns().size() - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }

        List<TextSegment> segments = compareTexts(originalText, correctedText, originalRuns, placeholders);

        for (TextSegment segment : segments) {
            XWPFRun run = paragraph.createRun();
            run.setText(segment.text, 0);

            if (segment.formatting != null) {
                segment.formatting.applyTo(run, segment.isChanged && !segment.isDeleted, segment.isDeleted);
            } else {
                defaultFormatting.applyTo(run, segment.isChanged && !segment.isDeleted, segment.isDeleted);
            }

            if (segment.isSuperscript) {
                run.setSubscript(VerticalAlign.SUPERSCRIPT);
            } else if (segment.isSubscript) {
                run.setSubscript(VerticalAlign.SUBSCRIPT);
            }
        }
    }

    private static List<TextSegment> compareTexts(String original, String corrected,
                                                   List<RunInfo> originalRuns,
                                                   List<PlaceholderInfo> placeholders) {
        // Build a character-to-formatting map from original runs
        Map<Integer, RunFormatting> charFormattingMap = new HashMap<>();
        int charIndex = 0;
        for (RunInfo runInfo : originalRuns) {
            for (int i = 0; i < runInfo.text.length(); i++) {
                charFormattingMap.put(charIndex++, runInfo.formatting);
            }
        }

        String[] origWords = original.split("(?<=\\s)|(?=\\s)|(?<=\\p{Punct})|(?=\\p{Punct})");
        String[] corrWords = corrected.split("(?<=\\s)|(?=\\s)|(?<=\\p{Punct})|(?=\\p{Punct})");

        List<DiffResult> diffs = computeWordDiff(origWords, corrWords);
        List<TextSegment> segments = new ArrayList<>();

        // Track position in original text to map formatting
        int origCharPos = 0;

        for (DiffResult diff : diffs) {
            boolean isChanged = diff.type != DiffType.UNCHANGED;
            boolean isDeleted = diff.type == DiffType.REMOVED;
            
            // Get formatting from the original position
            RunFormatting formatting = null;
            if (diff.type == DiffType.UNCHANGED || diff.type == DiffType.REMOVED) {
                // Use formatting from original position
                formatting = charFormattingMap.get(origCharPos);
                origCharPos += diff.text.length();
            } else {
                // For additions, try to use formatting from the context (previous character)
                formatting = charFormattingMap.get(Math.max(0, origCharPos - 1));
            }
            
            if (formatting == null) {
                formatting = originalRuns.isEmpty() ? new RunFormatting() : originalRuns.get(0).formatting;
            }
            
            segments.add(new TextSegment(diff.text, isChanged, isDeleted, false, false, formatting));
        }

        return processSuperSubscriptsInSegments(segments, placeholders);
    }

    private static List<DiffResult> computeWordDiff(String[] original, String[] corrected) {
        List<DiffResult> results = new ArrayList<>();
        int[][] dp = new int[original.length + 1][corrected.length + 1];

        for (int i = 1; i <= original.length; i++) {
            for (int j = 1; j <= corrected.length; j++) {
                if (original[i-1].equals(corrected[j-1])) {
                    dp[i][j] = dp[i-1][j-1] + 1;
                } else {
                    dp[i][j] = Math.max(dp[i-1][j], dp[i][j-1]);
                }
            }
        }

        int i = original.length;
        int j = corrected.length;

        while (i > 0 || j > 0) {
            if (i > 0 && j > 0 && original[i-1].equals(corrected[j-1])) {
                results.add(0, new DiffResult(original[i-1], DiffType.UNCHANGED));
                i--; j--;
            } else if (j > 0 && (i == 0 || dp[i][j-1] >= dp[i-1][j])) {
                results.add(0, new DiffResult(corrected[j-1], DiffType.ADDED));
                j--;
            } else if (i > 0) {
                results.add(0, new DiffResult(original[i-1], DiffType.REMOVED));
                i--;
            }
        }

        return results;
    }

    private static List<TextSegment> processSuperSubscriptsInSegments(List<TextSegment> segments,
                                                                       List<PlaceholderInfo> placeholders) {
        List<TextSegment> result = new ArrayList<>();

        for (TextSegment segment : segments) {
            String text = segment.text;
            int i = 0;
            StringBuilder currentText = new StringBuilder();

            while (i < text.length()) {
                char c = text.charAt(i);
                if (isSuperscriptOrSubscript(c)) {
                    if (currentText.length() > 0) {
                        result.add(new TextSegment(
                            currentText.toString(),
                            segment.isChanged,
                            segment.isDeleted,
                            false,
                            false,
                            segment.formatting
                        ));
                        currentText.setLength(0);
                    }

                    StringBuilder supsubGroup = new StringBuilder();
                    boolean isSuper = isSuperscriptChar(c);
                    while (i < text.length() && isSuperscriptOrSubscript(text.charAt(i))) {
                        supsubGroup.append(text.charAt(i));
                        i++;
                    }

                    // Do NOT mark Unicode super/sub as "changed" for coloring
                    result.add(new TextSegment(
                        supsubGroup.toString(),
                        false,
                        false,
                        isSuper,
                        !isSuper,
                        segment.formatting
                    ));
                } else {
                    currentText.append(c);
                    i++;
                }
            }

            if (currentText.length() > 0) {
                result.add(new TextSegment(
                    currentText.toString(),
                    segment.isChanged,
                    segment.isDeleted,
                    false,
                    false,
                    segment.formatting
                ));
            }
        }

        return result;
    }

    private static boolean isParagraphStyleToSkip(XWPFParagraph paragraph) {
        String styleId = paragraph.getStyleID();
        if (styleId != null && STYLES_TO_SKIP.contains(styleId.toUpperCase())) {
            return true;
        }

        for (XWPFRun run : paragraph.getRuns()) {
            try {
                CTR ctr = run.getCTR();
                if (ctr != null && ctr.isSetRPr()) {
                    CTString rStyle = ctr.getRPr().getRStyle();
                    if (rStyle != null && STYLES_TO_SKIP.contains(rStyle.getVal().toUpperCase())) {
                        return true;
                    }
                }
            } catch (Exception e) {
                // Ignore
            }
        }
        return false;
    }

    private static String callGrammarCheckApi(String text) throws IOException {
        URL url = new URL(API_URL);
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        connection.setRequestMethod("POST");
        connection.setRequestProperty("Content-Type", "application/json");
        connection.setDoOutput(true);

        String escapedText = text.replace("\\", "\\\\")
                                 .replace("\"", "\\\"")
                                 .replace("\n", "\\n")
                                 .replace("\r", "\\r");

        String payload = "{"
                + "\"model\":\"qwen2.5:3b\","
                + "\"messages\":["
                + "{\"role\":\"system\", \"content\":\"You are an expert copy editor. Your task is to review the provided text and return the corrected version of the text. ONLY fix grammatical errors, spelling mistakes, punctuation issues, and incorrect word usage in the main body text. DO NOT enhance, rewrite, or improve the sentence in any way.\\n\\n"
                + "CRITICAL RULES:\\n"
                + "- STRICTLY retain all existing quotes exactly as they are (straight or curved).\\n"
                + "- STRICTLY retain all brackets exactly as they are.\\n"
                + "- NEVER modify any placeholder text in the format «SUPSUB_N». These represent superscripts/subscripts that must remain exactly as is.\\n"
                + "- Do not alter the original tone, style, structure, or formatting intent.\\n"
                + "- Do not include any explanations, comments, or additional notes. Return ONLY the corrected text.\"},"
                + "{\"role\":\"user\", \"content\":\"" + escapedText + "\"}"
                + "],"
                + "\"temperature\":0.1"
                + "}";

        try (OutputStream os = connection.getOutputStream()) {
            os.write(payload.getBytes(StandardCharsets.UTF_8));
        }

        int responseCode = connection.getResponseCode();
        if (responseCode == 200) {
            try (BufferedReader br = new BufferedReader(new InputStreamReader(connection.getInputStream(), StandardCharsets.UTF_8))) {
                StringBuilder response = new StringBuilder();
                String line;
                while ((line = br.readLine()) != null) {
                    response.append(line.trim());
                }
                return extractCorrectedText(response.toString());
            }
        } else {
            throw new IOException("HTTP " + responseCode + " from API");
        }
    }

    private static String extractCorrectedText(String jsonResponse) {
        try {
            JSONObject jsonObject = new JSONObject(jsonResponse);
            JSONArray choices = jsonObject.getJSONArray("choices");
            if (choices.length() > 0) {
                return choices.getJSONObject(0).getJSONObject("message").getString("content");
            }
        } catch (Exception e) {
            System.err.println("Failed to parse JSON response: " + jsonResponse);
            e.printStackTrace();
        }
        return "No corrected text found.";
    }

    private static String writeToFile(XWPFDocument doc, String originalFileName) throws Exception {
        new File(OUTPUT_FOLDER).mkdirs();
        String correctedFileName = "T_" + originalFileName;
        try (FileOutputStream out = new FileOutputStream(OUTPUT_FOLDER + "/" + correctedFileName)) {
            doc.write(out);
        }
        return correctedFileName;
    }
}
