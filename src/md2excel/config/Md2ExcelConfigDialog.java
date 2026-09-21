package md2excel.config;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.FlowLayout;
import java.awt.Frame;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.KeyEvent;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.InvalidPathException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.ParseException;
import java.util.concurrent.atomic.AtomicReference;

import javax.swing.AbstractAction;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JRootPane;
import javax.swing.JSpinner;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.SpinnerNumberModel;
import javax.swing.SwingUtilities;
import javax.swing.WindowConstants;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;

import org.apache.poi.ss.usermodel.VerticalAlignment;

public final class Md2ExcelConfigDialog {

    private static final String DEFAULT_FONT_NAME = "游ゴシック";

    private static final int DEFAULT_H1_FONT_SIZE = 16;
    private static final int DEFAULT_H2_FONT_SIZE = 14;
    private static final int DEFAULT_H3_FONT_SIZE = 12;
    private static final int DEFAULT_NORMAL_FONT_SIZE = 11;
    private static final int DEFAULT_TOTAL_COLUMN_COUNT = 40;

    private static final String[] FONT_CANDIDATES = { "游ゴシック", "Yu Gothic UI", "ＭＳ Ｐゴシック", "ＭＳ ゴシック", "Meiryo",
            "Meiryo UI" };

    private static final VerticalAlignmentOption[] ALIGNMENT_OPTIONS = {
            new VerticalAlignmentOption("上揃え", VerticalAlignment.TOP),
            new VerticalAlignmentOption("上下中央揃え", VerticalAlignment.CENTER),
            new VerticalAlignmentOption("下揃え", VerticalAlignment.BOTTOM) };

    private static final VerticalAlignmentOption DEFAULT_ALIGNMENT_OPTION = ALIGNMENT_OPTIONS[2];

    private Md2ExcelConfigDialog() {
    }

    /**
     * 設定フォームを表示する。
     *
     * @return 入力された設定。キャンセル時はnull
     */
    public static Md2ExcelConfig load() {
        if (SwingUtilities.isEventDispatchThread()) {
            return new ConfigForm().showDialog();
        }

        final AtomicReference<Md2ExcelConfig> result = new AtomicReference<Md2ExcelConfig>();

        try {
            SwingUtilities.invokeAndWait(new Runnable() {
                @Override
                public void run() {
                    result.set(new ConfigForm().showDialog());
                }
            });
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();

            throw new IllegalStateException("Configuration dialog was interrupted", e);

        } catch (InvocationTargetException e) {
            Throwable cause = e.getCause();

            if (cause instanceof RuntimeException) {
                throw (RuntimeException) cause;
            }

            if (cause instanceof Error) {
                throw (Error) cause;
            }

            throw new IllegalStateException("Failed to show configuration dialog", cause);
        }

        return result.get();
    }

    /**
     * 入力パスと同じディレクトリに、 拡張子だけを変更した出力パスを作る。
     */
    static Path replaceExtension(Path path, String newExtension) {

        if (path == null) {
            throw new IllegalArgumentException("path must not be null");
        }

        if (newExtension == null || newExtension.trim().isEmpty()) {

            throw new IllegalArgumentException("newExtension must not be empty");
        }

        Path fileNamePath = path.getFileName();

        if (fileNamePath == null) {
            throw new IllegalArgumentException("path must have a file name: " + path);
        }

        String fileName = fileNamePath.toString();

        int dot = fileName.lastIndexOf('.');

        String outputFileName = dot > 0 ? fileName.substring(0, dot) + newExtension : fileName + newExtension;

        Path parent = path.getParent();

        return parent == null ? Paths.get(outputFileName) : parent.resolve(outputFileName);
    }

    private static final class ConfigForm {

        private final JTextField inputPathField = new JTextField(36);

        private final JTextField outputPathField = new JTextField(36);

        private final JButton browseButton = new JButton("参照...");

        private final JSpinner totalColumnCountSpinner = createIntegerSpinner(DEFAULT_TOTAL_COLUMN_COUNT,
                MdSheetSettings.MIN_TOTAL_COLUMN_COUNT, MdSheetSettings.MAX_TOTAL_COLUMN_COUNT);

        private final JComboBox<String> fontComboBox = new JComboBox<String>(FONT_CANDIDATES);

        private final JComboBox<VerticalAlignmentOption> alignmentComboBox = new JComboBox<VerticalAlignmentOption>(
                ALIGNMENT_OPTIONS);

        private final JSpinner h1SizeSpinner = createIntegerSpinner(DEFAULT_H1_FONT_SIZE, MdFontSettings.MIN_FONT_SIZE,
                MdFontSettings.MAX_FONT_SIZE);

        private final JSpinner h2SizeSpinner = createIntegerSpinner(DEFAULT_H2_FONT_SIZE, MdFontSettings.MIN_FONT_SIZE,
                MdFontSettings.MAX_FONT_SIZE);

        private final JSpinner h3SizeSpinner = createIntegerSpinner(DEFAULT_H3_FONT_SIZE, MdFontSettings.MIN_FONT_SIZE,
                MdFontSettings.MAX_FONT_SIZE);

        private final JSpinner normalSizeSpinner = createIntegerSpinner(DEFAULT_NORMAL_FONT_SIZE,
                MdFontSettings.MIN_FONT_SIZE, MdFontSettings.MAX_FONT_SIZE);

        private final JLabel errorLabel = new JLabel(" ");

        private final JDialog dialog;

        private Md2ExcelConfig result;

        ConfigForm() {
            dialog = new JDialog((Frame) null, "Markdown → Excel 設定", true);

            configureComponents();
            configureDialog();
        }

        Md2ExcelConfig showDialog() {
            dialog.setVisible(true);
            return result;
        }

        private void configureComponents() {
            outputPathField.setEditable(false);
            outputPathField.setFocusable(false);

            fontComboBox.setEditable(false);
            fontComboBox.setSelectedItem(DEFAULT_FONT_NAME);

            alignmentComboBox.setSelectedItem(DEFAULT_ALIGNMENT_OPTION);

            errorLabel.setForeground(new Color(180, 0, 0));

            browseButton.addActionListener(new java.awt.event.ActionListener() {
                @Override
                public void actionPerformed(ActionEvent event) {

                    chooseInputFile();
                }
            });

            inputPathField.getDocument().addDocumentListener(new DocumentListener() {
                @Override
                public void insertUpdate(DocumentEvent event) {

                    updateOutputPathPreview();
                }

                @Override
                public void removeUpdate(DocumentEvent event) {

                    updateOutputPathPreview();
                }

                @Override
                public void changedUpdate(DocumentEvent event) {

                    updateOutputPathPreview();
                }
            });
        }

        private void configureDialog() {
            dialog.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);

            JPanel content = new JPanel(new BorderLayout(0, 12));

            content.setBorder(BorderFactory.createEmptyBorder(12, 12, 12, 12));

            content.add(createFormPanel(), BorderLayout.CENTER);

            content.add(createButtonPanel(), BorderLayout.SOUTH);

            dialog.setContentPane(content);

            bindEscapeKey();

            dialog.pack();
            dialog.setResizable(false);
            dialog.setLocationRelativeTo(null);
        }

        private JPanel createFormPanel() {
            JPanel panel = new JPanel(new GridBagLayout());

            GridBagConstraints constraints = new GridBagConstraints();

            constraints.insets = new Insets(4, 4, 4, 4);

            constraints.anchor = GridBagConstraints.WEST;

            constraints.fill = GridBagConstraints.HORIZONTAL;

            constraints.weightx = 1.0;

            int row = 0;

            JPanel inputPanel = new JPanel(new BorderLayout(6, 0));

            inputPanel.add(inputPathField, BorderLayout.CENTER);

            inputPanel.add(browseButton, BorderLayout.EAST);

            addFormRow(panel, constraints, row++, "Markdownファイル:", inputPanel);

            addFormRow(panel, constraints, row++, "出力ファイル:", outputPathField);

            addFormRow(panel, constraints, row++, "シート全体の列数 (" + MdSheetSettings.MIN_TOTAL_COLUMN_COUNT + "～"
                    + MdSheetSettings.MAX_TOTAL_COLUMN_COUNT + "):", totalColumnCountSpinner);

            addFormRow(panel, constraints, row++, "フォント:", fontComboBox);

            addFormRow(panel, constraints, row++, "セルの縦位置:", alignmentComboBox);

            addFormRow(panel, constraints, row++, "# 見出しサイズ:", h1SizeSpinner);

            addFormRow(panel, constraints, row++, "## 見出しサイズ:", h2SizeSpinner);

            addFormRow(panel, constraints, row++, "### 見出しサイズ:", h3SizeSpinner);

            addFormRow(panel, constraints, row++, "通常テキストサイズ:", normalSizeSpinner);

            GridBagConstraints errorConstraints = new GridBagConstraints();

            errorConstraints.gridx = 0;
            errorConstraints.gridy = row;
            errorConstraints.gridwidth = 2;
            errorConstraints.weightx = 1.0;
            errorConstraints.fill = GridBagConstraints.HORIZONTAL;

            errorConstraints.insets = new Insets(8, 4, 0, 4);

            panel.add(errorLabel, errorConstraints);

            return panel;
        }

        private JPanel createButtonPanel() {
            JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT));

            JButton okButton = new JButton("OK");

            JButton cancelButton = new JButton("キャンセル");

            okButton.addActionListener(new java.awt.event.ActionListener() {
                @Override
                public void actionPerformed(ActionEvent event) {

                    submit();
                }
            });

            cancelButton.addActionListener(new java.awt.event.ActionListener() {
                @Override
                public void actionPerformed(ActionEvent event) {

                    cancel();
                }
            });

            dialog.getRootPane().setDefaultButton(okButton);

            panel.add(okButton);
            panel.add(cancelButton);

            return panel;
        }

        private void chooseInputFile() {
            JFileChooser fileChooser = new JFileChooser();

            fileChooser.setDialogTitle("Markdown ファイルを選択してください");

            fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);

            setInitialFileChooserLocation(fileChooser);

            int chooserResult = fileChooser.showOpenDialog(dialog);

            if (chooserResult != JFileChooser.APPROVE_OPTION) {

                // ファイル選択だけをキャンセルし、
                // 設定フォームには戻る。
                return;
            }

            Path selectedPath = fileChooser.getSelectedFile().toPath().toAbsolutePath();

            inputPathField.setText(selectedPath.toString());
        }

        private void setInitialFileChooserLocation(JFileChooser fileChooser) {

            String currentText = inputPathField.getText().trim();

            if (currentText.isEmpty()) {
                return;
            }

            try {
                Path currentPath = Paths.get(currentText).toAbsolutePath();

                Path directory = currentPath.getParent();

                if (directory != null) {
                    fileChooser.setCurrentDirectory(directory.toFile());
                }

                fileChooser.setSelectedFile(currentPath.toFile());

            } catch (InvalidPathException ignored) {
                // 入力途中の不正なパスは無視する。
            }
        }

        private void updateOutputPathPreview() {
            String inputText = inputPathField.getText().trim();

            if (inputText.isEmpty()) {
                outputPathField.setText("");
                return;
            }

            try {
                Path inputPath = Paths.get(inputText).toAbsolutePath();

                Path outputPath = replaceExtension(inputPath, ".xlsx");

                outputPathField.setText(outputPath.toString());

            } catch (InvalidPathException e) {
                outputPathField.setText("");
            }
        }

        private void submit() {
            errorLabel.setText(" ");

            try {
                commitSpinnerEditors();

                String inputText = inputPathField.getText().trim();

                if (inputText.isEmpty()) {
                    showError("Markdownファイルを選択してください。");
                    return;
                }

                Path inputPath = Paths.get(inputText).toAbsolutePath();

                Path outputPath = replaceExtension(inputPath, ".xlsx");

                Object selectedFont = fontComboBox.getSelectedItem();

                if (selectedFont == null || selectedFont.toString().trim().isEmpty()) {

                    showError("フォントを選択してください。");
                    return;
                }

                VerticalAlignmentOption alignmentOption = (VerticalAlignmentOption) alignmentComboBox.getSelectedItem();

                if (alignmentOption == null) {
                    showError("セルの縦位置を選択してください。");
                    return;
                }

                MdFontSettings fontSettings = new MdFontSettings(selectedFont.toString(),
                        getSpinnerValue(h1SizeSpinner), getSpinnerValue(h2SizeSpinner), getSpinnerValue(h3SizeSpinner),
                        getSpinnerValue(normalSizeSpinner));

                MdSheetSettings sheetSettings = new MdSheetSettings(getSpinnerValue(totalColumnCountSpinner),
                        alignmentOption.getAlignment());

                result = new Md2ExcelConfig(inputPath, outputPath, fontSettings, sheetSettings);

                dialog.dispose();

            } catch (InvalidPathException e) {
                showError("入力ファイルのパスが正しくありません。");

            } catch (ParseException e) {
                showError("数値項目には指定範囲内の整数を" + "入力してください。");

            } catch (IllegalArgumentException e) {
                showError(e.getMessage());
            }
        }

        private void commitSpinnerEditors() throws ParseException {

            totalColumnCountSpinner.commitEdit();
            h1SizeSpinner.commitEdit();
            h2SizeSpinner.commitEdit();
            h3SizeSpinner.commitEdit();
            normalSizeSpinner.commitEdit();
        }

        private void cancel() {
            result = null;
            dialog.dispose();
        }

        private void bindEscapeKey() {
            JRootPane rootPane = dialog.getRootPane();

            String actionKey = "cancel-form";

            rootPane.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0),
                    actionKey);

            rootPane.getActionMap().put(actionKey, new AbstractAction() {
                private static final long serialVersionUID = 1L;

                @Override
                public void actionPerformed(ActionEvent event) {

                    cancel();
                }
            });
        }

        private void showError(String message) {
            String safeMessage = message == null || message.trim().isEmpty() ? "入力内容を確認してください。" : message;

            errorLabel.setText("入力エラー: " + safeMessage);
        }
    }

    private static void addFormRow(JPanel panel, GridBagConstraints baseConstraints, int row, String labelText,
            JComponent component) {

        GridBagConstraints labelConstraints = (GridBagConstraints) baseConstraints.clone();

        labelConstraints.gridx = 0;
        labelConstraints.gridy = row;
        labelConstraints.weightx = 0.0;
        labelConstraints.fill = GridBagConstraints.NONE;

        panel.add(new JLabel(labelText), labelConstraints);

        GridBagConstraints inputConstraints = (GridBagConstraints) baseConstraints.clone();

        inputConstraints.gridx = 1;
        inputConstraints.gridy = row;
        inputConstraints.weightx = 1.0;
        inputConstraints.fill = GridBagConstraints.HORIZONTAL;

        panel.add(component, inputConstraints);
    }

    private static JSpinner createIntegerSpinner(int defaultValue, int minValue, int maxValue) {

        JSpinner spinner = new JSpinner(new SpinnerNumberModel(defaultValue, minValue, maxValue, 1));

        JSpinner.NumberEditor editor = new JSpinner.NumberEditor(spinner, "0");

        editor.getTextField().setColumns(6);

        spinner.setEditor(editor);

        return spinner;
    }

    private static int getSpinnerValue(JSpinner spinner) {

        Object value = spinner.getValue();

        if (!(value instanceof Number)) {
            throw new IllegalArgumentException("Spinner value must be numeric");
        }

        return ((Number) value).intValue();
    }

    private static final class VerticalAlignmentOption {

        private final String label;
        private final VerticalAlignment alignment;

        VerticalAlignmentOption(String label, VerticalAlignment alignment) {

            this.label = label;
            this.alignment = alignment;
        }

        VerticalAlignment getAlignment() {
            return alignment;
        }

        @Override
        public String toString() {
            return label;
        }
    }
}