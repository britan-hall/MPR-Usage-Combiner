package com.britan.mprcombiner;

import com.formdev.flatlaf.FlatClientProperties;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.border.EmptyBorder;
import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.Window;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Locale;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicReference;

/** Modal output path and post-run options before combine starts. */
final class CombineOutputDialog {

    record Result(
            Path explicitOutput,
            Path outputDirWhenAuto,
            boolean openWorkbookAfter,
            boolean copyCopilotPrompt
    ) {
    }

    private CombineOutputDialog() {
    }

    static Optional<Result> show(
            Window parent,
            Path combineFromDir,
            String rememberedOutPath,
            String rememberedFolderPath,
            boolean rememberedOpenWorkbook,
            boolean rememberedCopyCopilot
    ) {
        JDialog dlg = new JDialog(parent, "Output", java.awt.Dialog.ModalityType.APPLICATION_MODAL);
        dlg.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);

        JTextField outputPathField = new JTextField(rememberedOutPath);
        styleField(outputPathField);
        outputPathField.putClientProperty(FlatClientProperties.PLACEHOLDER_TEXT, "Optional .xlsx or .csv path");

        JTextField outputFolderField = new JTextField(rememberedFolderPath);
        styleField(outputFolderField);
        outputFolderField.putClientProperty(FlatClientProperties.PLACEHOLDER_TEXT, "Optional folder for auto-named file");

        JButton browseOut = UiTheme.secondaryButton("Save as…");
        browseOut.addActionListener(e -> MPRCombinerGui.pickOutputFile(dlg, outputPathField, combineFromDir));
        JButton browseFolder = UiTheme.secondaryButton("Browse…");
        browseFolder.addActionListener(e -> MPRCombinerGui.pickDirectory(
                dlg, outputFolderField, combineFromDir, "Output folder"));

        JPanel form = new JPanel(new GridBagLayout());
        form.setOpaque(false);
        GridBagConstraints g = new GridBagConstraints();
        g.gridx = 0;
        g.gridy = 0;
        g.anchor = GridBagConstraints.FIRST_LINE_START;
        g.insets = new Insets(0, 0, 8, 12);
        form.add(label("Output file"), g);
        g.gridx = 1;
        g.weightx = 1;
        g.fill = GridBagConstraints.HORIZONTAL;
        g.insets = new Insets(0, 0, 8, 10);
        form.add(outputPathField, g);
        g.gridx = 2;
        g.weightx = 0;
        g.fill = GridBagConstraints.NONE;
        g.insets = new Insets(0, 0, 8, 0);
        form.add(browseOut, g);

        g.gridx = 0;
        g.gridy = 1;
        g.insets = new Insets(0, 0, 8, 12);
        form.add(label("Output folder"), g);
        g.gridx = 1;
        g.weightx = 1;
        g.fill = GridBagConstraints.HORIZONTAL;
        g.insets = new Insets(0, 0, 8, 10);
        form.add(outputFolderField, g);
        g.gridx = 2;
        g.weightx = 0;
        g.fill = GridBagConstraints.NONE;
        g.insets = new Insets(0, 0, 8, 0);
        form.add(browseFolder, g);

        JCheckBox openWorkbook = new JCheckBox("Open workbook when finished");
        openWorkbook.setOpaque(false);
        openWorkbook.setForeground(UiTheme.TEXT_SECONDARY);
        openWorkbook.setSelected(rememberedOpenWorkbook);

        JCheckBox copyCopilot = new JCheckBox("Copy Copilot prompt to clipboard");
        copyCopilot.setOpaque(false);
        copyCopilot.setForeground(UiTheme.TEXT_SECONDARY);
        copyCopilot.setSelected(rememberedCopyCopilot);

        g = new GridBagConstraints();
        g.gridx = 0;
        g.gridy = 2;
        g.gridwidth = 3;
        g.anchor = GridBagConstraints.WEST;
        g.insets = new Insets(4, 0, 4, 0);
        form.add(openWorkbook, g);
        g.gridy = 3;
        g.insets = new Insets(0, 0, 0, 0);
        form.add(copyCopilot, g);

        AtomicReference<Optional<Result>> outcome = new AtomicReference<>(Optional.empty());

        JButton cancel = UiTheme.secondaryButton("Cancel");
        cancel.addActionListener(e -> {
            outcome.set(Optional.empty());
            dlg.dispose();
        });

        JButton go = UiTheme.primaryButton("Combine");
        go.addActionListener(e -> {
            Optional<Result> parsed = parseResult(
                    dlg,
                    outputPathField.getText(),
                    outputFolderField.getText(),
                    openWorkbook.isSelected(),
                    copyCopilot.isSelected());
            if (parsed.isEmpty()) {
                return;
            }
            outcome.set(parsed);
            dlg.dispose();
        });
        dlg.getRootPane().setDefaultButton(go);

        JPanel buttons = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 0));
        buttons.setOpaque(false);
        buttons.setBorder(new EmptyBorder(16, 0, 0, 0));
        buttons.add(cancel);
        buttons.add(go);

        JPanel shell = new JPanel(new BorderLayout(0, 0));
        shell.setBorder(new EmptyBorder(20, 24, 20, 24));
        shell.setBackground(UiTheme.CARD_BG);
        shell.add(form, BorderLayout.CENTER);
        shell.add(buttons, BorderLayout.SOUTH);

        dlg.setContentPane(shell);
        dlg.pack();
        dlg.setMinimumSize(new Dimension(520, 0));
        dlg.setLocationRelativeTo(parent);
        dlg.setVisible(true);
        return outcome.get();
    }

    private static Optional<Result> parseResult(
            Component dialogParent,
            String outRaw,
            String folderRaw,
            boolean openWorkbook,
            boolean copyCopilot
    ) {
        String outTrim = outRaw.trim();
        final Path explicitOutput;
        if (outTrim.isEmpty()) {
            explicitOutput = null;
        } else {
            Path p = Paths.get(outTrim).toAbsolutePath().normalize();
            String name = p.getFileName() != null ? p.getFileName().toString().toLowerCase(Locale.ROOT) : "";
            if (!name.endsWith(".xlsx") && !name.endsWith(".csv")) {
                JOptionPane.showMessageDialog(
                        dialogParent,
                        "Output file should end with .xlsx or .csv,\nor leave the field blank for the automatic name.",
                        "Output",
                        JOptionPane.WARNING_MESSAGE);
                return Optional.empty();
            }
            explicitOutput = p;
        }

        final Path outputDirWhenAuto;
        if (explicitOutput != null) {
            outputDirWhenAuto = null;
        } else {
            String f = folderRaw.trim();
            if (f.isEmpty()) {
                outputDirWhenAuto = null;
            } else {
                Path d = Paths.get(f).toAbsolutePath().normalize();
                if (!Files.isDirectory(d)) {
                    JOptionPane.showMessageDialog(
                            dialogParent,
                            "Output folder must be an existing directory:\n" + d,
                            "Output",
                            JOptionPane.WARNING_MESSAGE);
                    return Optional.empty();
                }
                outputDirWhenAuto = d;
            }
        }
        return Optional.of(new Result(explicitOutput, outputDirWhenAuto, openWorkbook, copyCopilot));
    }

    private static JLabel label(String text) {
        JLabel l = new JLabel(text);
        l.setFont(l.getFont().deriveFont(Font.PLAIN, 13f));
        l.setForeground(UiTheme.TEXT_SECONDARY);
        return l;
    }

    private static void styleField(JTextField field) {
        field.setFont(field.getFont().deriveFont(Font.PLAIN, 13f));
        field.putClientProperty(FlatClientProperties.STYLE, "arc: 6");
        field.setMargin(new Insets(8, 10, 8, 10));
    }
}
