package com.britan.mprcombiner;

import com.formdev.flatlaf.FlatClientProperties;
import com.formdev.flatlaf.FlatLightLaf;

import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.ScrollPaneConstants;
import javax.swing.JSplitPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.ImageIcon;
import javax.swing.JTree;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.Timer;
import javax.swing.UIManager;
import javax.swing.event.TreeExpansionEvent;
import javax.swing.event.TreeExpansionListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreePath;
import javax.swing.tree.TreeSelectionModel;
import java.awt.AlphaComposite;
import java.awt.BasicStroke;
import java.awt.BorderLayout;
import java.awt.CardLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.GradientPaint;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.FlowLayout;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Image;
import java.awt.Insets;
import java.awt.RenderingHints;
import java.io.OutputStream;
import java.io.PrintStream;
import java.io.IOException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.Set;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;
import java.util.stream.Stream;

/** Desktop UI for combining MPR Usage workbooks from a local or synced folder. */
public final class MPRCombinerGui {

    private static final Object TREE_PLACEHOLDER = new Object();

    private static final Set<String> EXCEL_EXT = Set.of("xlsx", "xlsm", "xls");

    private static final int TREE_ROW_HEIGHT = 28;
    private static final int CARD_PAD = 20;
    private static final int SOURCE_LABEL_WIDTH = 116;
    private static final int SOURCE_FIELD_GAP = 10;
    private static final int BRAND_LOGO_TARGET_HEIGHT = 46;
    private static final String DEFAULT_SHEET_NAME = "Usage";
    private static final int DEFAULT_HEADER_ROW = 1;
    /** Default wall-clock limit for one combine run (stuck reads when cloud files are not fully synced). */
    private static final int DEFAULT_COMBINE_TIMEOUT_MINUTES = 15;

    private static int combineTimeoutMinutes() {
        String p = System.getProperty("mpr.combiner.timeoutMinutes");
        if (p != null) {
            try {
                int v = Integer.parseInt(p.trim());
                if (v >= 1 && v <= 240) {
                    return v;
                }
            } catch (NumberFormatException ignored) {
                // fall through
            }
        }
        return DEFAULT_COMBINE_TIMEOUT_MINUTES;
    }

    /** Remembered across Combine runs in this session (output dialog). */
    private String rememberedOutputPath = "";
    private String rememberedOutputFolder = "";
    private boolean rememberedOpenWorkbook = false;
    private boolean rememberedCopyCopilot = true;

    /**
     * Optional company mark: {@code /branding/logo.png} or {@code /branding/logo.jpg} on the classpath
     * (e.g. {@code src/main/resources/branding/logo.png} in this project).
     */
    private static Optional<ImageIcon> loadBrandingLogo(int targetHeightPx) {
        String[] resourcePaths = {"/branding/logo.png", "/branding/logo.jpg"};
        for (String path : resourcePaths) {
            URL url = MPRCombinerGui.class.getResource(path);
            if (url == null) {
                continue;
            }
            ImageIcon original = new ImageIcon(url);
            int ih = original.getIconHeight();
            int iw = original.getIconWidth();
            if (ih <= 0 || iw <= 0) {
                continue;
            }
            double scale = (double) targetHeightPx / (double) ih;
            scale = Math.min(scale, 3.0);
            int nw = Math.max(1, (int) Math.round(iw * scale));
            int nh = Math.max(1, (int) Math.round(ih * scale));
            Image scaled = original.getImage().getScaledInstance(nw, nh, Image.SCALE_SMOOTH);
            return Optional.of(new ImageIcon(scaled));
        }
        return Optional.empty();
    }

    private MPRCombinerGui() {
    }

    /** Dotfiles and OS-hidden entries are omitted from the library tree. */
    private static boolean isHiddenPath(Path path) {
        Path name = path.getFileName();
        if (name == null) {
            return false;
        }
        String s = name.toString();
        if (!s.isEmpty() && s.charAt(0) == '.') {
            return true;
        }
        try {
            return Files.isHidden(path);
        } catch (IOException e) {
            return false;
        }
    }

    private static void configureFileChooser(JFileChooser chooser) {
        chooser.setFileHidingEnabled(true);
    }

    private static void styleTextField(JTextField field) {
        field.setMargin(new Insets(8, 12, 8, 12));
    }

    public static void launch() {
        SwingUtilities.invokeLater(() -> {
            try {
                FlatLightLaf.setup();
                UIManager.put("Button.arc", 8);
                UIManager.put("Component.arc", 8);
                UIManager.put("TextComponent.arc", 8);
                UIManager.put("ScrollBar.thumbArc", 999);
                UIManager.put("ScrollBar.width", 10);
                UIManager.put("Table.showHorizontalLines", true);
                if (System.getProperty("os.name", "").toLowerCase(Locale.ROOT).contains("mac")) {
                    UIManager.put("TitlePane.unifiedBackground", Boolean.TRUE);
                }
            } catch (Exception ignored) {
                try {
                    UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
                } catch (Exception ignored2) {
                    // default
                }
            }
            new MPRCombinerGui().open();
        });
    }

    private void open() {
        JFrame frame = new JFrame("MPR Combiner");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        JTextField inputDirField = new JTextField();
        styleTextField(inputDirField);
        inputDirField.putClientProperty(FlatClientProperties.PLACEHOLDER_TEXT, "Path to folder containing Excel workbooks");
        JCheckBox recursiveBox = new JCheckBox("Include subfolders");
        recursiveBox.setSelected(true);
        recursiveBox.setOpaque(false);
        recursiveBox.setForeground(UiTheme.TEXT_SECONDARY);
        recursiveBox.setFont(recursiveBox.getFont().deriveFont(Font.PLAIN, 13f));
        JTextArea logArea = new JTextArea(1, 1);
        logArea.setEditable(false);
        logArea.setLineWrap(true);
        logArea.setWrapStyleWord(true);
        logArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        logArea.setMargin(new Insets(12, 14, 12, 14));
        logArea.setBackground(UiTheme.LOG_SURFACE);
        logArea.setForeground(UiTheme.TEXT_PRIMARY);
        logArea.putClientProperty(FlatClientProperties.STYLE, "arc: 8");

        DefaultMutableTreeNode treeRoot = new DefaultMutableTreeNode("");
        DefaultTreeModel fileTreeModel = new DefaultTreeModel(treeRoot);
        JTree fileTree = new JTree(fileTreeModel);
        fileTree.setRootVisible(false);
        fileTree.setShowsRootHandles(false);
        fileTree.setRowHeight(TREE_ROW_HEIGHT);
        fileTree.setFont(fileTree.getFont().deriveFont(Font.PLAIN, 13f));
        fileTree.setBackground(UiTheme.CARD_BG);
        fileTree.getSelectionModel().setSelectionMode(TreeSelectionModel.SINGLE_TREE_SELECTION);
        fileTree.setBorder(BorderFactory.createEmptyBorder(4, 8, 4, 8));
        JScrollPane fileTreeScroll = new JScrollPane(fileTree);
        fileTreeScroll.setBorder(BorderFactory.createEmptyBorder());
        fileTreeScroll.getViewport().setBackground(UiTheme.CARD_BG);
        fileTreeScroll.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        fileTreeScroll.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED);
        fileTreeScroll.setMinimumSize(new Dimension(240, 120));
        CardLayout treeCards = new CardLayout();
        JPanel treePanel = new JPanel(treeCards);
        treePanel.setOpaque(true);
        treePanel.setBackground(UiTheme.CARD_BG);
        treePanel.setBorder(BorderFactory.createLineBorder(UiTheme.LOG_BORDER, 1, true));
        JPanel emptyTreeState = new JPanel(new GridBagLayout());
        emptyTreeState.setOpaque(true);
        emptyTreeState.setBackground(UiTheme.CARD_BG);
        JButton openLibraryBtn = UiTheme.secondaryButton("Open library");
        openLibraryBtn.addActionListener(e ->
                pickSyncedLibraryRoot(frame, fileTree, fileTreeModel, inputDirField, treePanel));
        GridBagConstraints emptyTreeGc = new GridBagConstraints();
        emptyTreeGc.anchor = GridBagConstraints.CENTER;
        emptyTreeState.add(openLibraryBtn, emptyTreeGc);
        treePanel.add(emptyTreeState, "empty");
        treePanel.add(fileTreeScroll, "tree");
        treeCards.show(treePanel, "empty");
        fileTree.addTreeExpansionListener(new LocalFolderTreeExpansionListener(fileTree, fileTreeModel));
        fileTree.addTreeSelectionListener(e -> {
            TreePath path = fileTree.getSelectionPath();
            if (path == null) {
                return;
            }
            DefaultMutableTreeNode n = (DefaultMutableTreeNode) path.getLastPathComponent();
            Object u = n.getUserObject();
            if (u instanceof FileTreePayload p) {
                if (p.directory) {
                    inputDirField.setText(p.path.toAbsolutePath().normalize().toString());
                } else {
                    Path parent = p.path.getParent();
                    if (parent != null) {
                        inputDirField.setText(parent.toAbsolutePath().normalize().toString());
                    }
                }
            }
        });

        JButton browseInBtn = UiTheme.compactSecondaryButton("Browse…");
        browseInBtn.addActionListener(e -> {
            if (!pickDirectory(frame, inputDirField, null, "Folder containing Excel workbooks")) {
                return;
            }
            String chosen = inputDirField.getText().trim();
            if (!chosen.isEmpty()) {
                Path root = Paths.get(chosen).toAbsolutePath().normalize();
                if (Files.isDirectory(root)) {
                    loadLibraryTree(fileTree, fileTreeModel, inputDirField, root, treePanel);
                }
            }
        });

        JPanel sourceBody = new JPanel(new GridBagLayout());
        sourceBody.setOpaque(false);
        GridBagConstraints gc = new GridBagConstraints();
        gc.gridx = 0;
        gc.gridy = 0;
        gc.gridwidth = 3;
        gc.weightx = 1;
        gc.weighty = 1;
        gc.fill = GridBagConstraints.BOTH;
        gc.insets = new Insets(0, 0, 12, 0);
        sourceBody.add(treePanel, gc);

        gc.gridy = 1;
        gc.gridwidth = 1;
        gc.weightx = 0;
        gc.weighty = 0;
        gc.fill = GridBagConstraints.NONE;
        gc.insets = new Insets(0, 0, 0, SOURCE_FIELD_GAP);
        gc.anchor = GridBagConstraints.BASELINE;
        sourceBody.add(sourceFieldLabel("Combine from"), gc);

        gc.gridx = 1;
        gc.weightx = 1;
        gc.fill = GridBagConstraints.HORIZONTAL;
        gc.insets = new Insets(0, 0, 0, SOURCE_FIELD_GAP);
        sourceBody.add(inputDirField, gc);

        gc.gridx = 2;
        gc.weightx = 0;
        gc.fill = GridBagConstraints.NONE;
        gc.insets = new Insets(0, 0, 0, 0);
        gc.anchor = GridBagConstraints.BASELINE;
        sourceBody.add(browseInBtn, gc);

        StatusIndicator statusIndicator = new StatusIndicator();

        JPanel sourceCard = cardPanel("Source", sourceBody, recursiveBox);

        JPanel logCard = new JPanel(new BorderLayout(0, 10));
        logCard.setOpaque(true);
        logCard.setBackground(UiTheme.CARD_BG);
        logCard.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(UiTheme.CARD_BORDER, 1, true),
                BorderFactory.createEmptyBorder(14, 18, 16, 18)));
        JLabel logTitle = new JLabel("Activity");
        logTitle.setFont(logTitle.getFont().deriveFont(Font.BOLD, 13f));
        logTitle.setForeground(UiTheme.TEXT_PRIMARY);
        logCard.add(logTitle, BorderLayout.NORTH);
        JScrollPane logScroll = new JScrollPane(logArea);
        logScroll.setBorder(BorderFactory.createLineBorder(UiTheme.LOG_BORDER, 1, true));
        logScroll.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        logScroll.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED);
        logScroll.getViewport().setBackground(UiTheme.LOG_SURFACE);
        logCard.add(logScroll, BorderLayout.CENTER);

        JButton runButton = UiTheme.primaryButton("Combine");
        Dimension runButtonSize = runButton.getPreferredSize();
        runButton.setMinimumSize(runButtonSize);
        runButton.setMaximumSize(runButtonSize);

        JPanel actionBar = new JPanel(new GridBagLayout());
        actionBar.setOpaque(false);
        actionBar.setBorder(BorderFactory.createEmptyBorder(14, 0, 0, 0));
        GridBagConstraints actionGc = new GridBagConstraints();
        actionGc.gridy = 0;
        actionGc.gridx = 0;
        actionGc.anchor = GridBagConstraints.WEST;
        actionBar.add(statusIndicator, actionGc);
        actionGc.gridx = 1;
        actionGc.weightx = 1;
        actionGc.fill = GridBagConstraints.HORIZONTAL;
        JPanel actionFiller = new JPanel();
        actionFiller.setOpaque(false);
        actionBar.add(actionFiller, actionGc);
        actionGc.gridx = 2;
        actionGc.weightx = 0;
        actionGc.fill = GridBagConstraints.NONE;
        actionGc.anchor = GridBagConstraints.EAST;
        actionBar.add(runButton, actionGc);

        JPanel workColumn = new JPanel(new BorderLayout(0, 0));
        workColumn.setOpaque(false);
        workColumn.setBorder(BorderFactory.createEmptyBorder(12, 28, 16, 12));
        workColumn.add(sourceCard, BorderLayout.CENTER);
        workColumn.add(actionBar, BorderLayout.SOUTH);

        JPanel logColumn = new JPanel(new BorderLayout());
        logColumn.setOpaque(false);
        logColumn.setBorder(BorderFactory.createEmptyBorder(12, 12, 16, 28));
        logColumn.setMinimumSize(new Dimension(280, 0));
        logColumn.add(logCard, BorderLayout.CENTER);

        JSplitPane split = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, workColumn, logColumn);
        split.setResizeWeight(0.58);
        split.setOpaque(false);
        split.setBorder(BorderFactory.createEmptyBorder(0, 0, 0, 0));
        split.setDividerSize(8);
        split.setOneTouchExpandable(true);
        split.setContinuousLayout(true);

        JPanel root = new JPanel(new BorderLayout());
        root.setOpaque(true);
        root.setBackground(UiTheme.BG_APP);
        root.add(new BrandHeader(loadBrandingLogo(BRAND_LOGO_TARGET_HEIGHT)), BorderLayout.NORTH);
        root.add(split, BorderLayout.CENTER);

        frame.add(root);
        frame.getRootPane().setDefaultButton(runButton);

        runButton.addActionListener(e -> {
            if (!validateCombineSource(frame, inputDirField)) {
                return;
            }
            Path inputDir = Paths.get(inputDirField.getText().trim()).toAbsolutePath().normalize();
            Optional<CombineOutputDialog.Result> choice = CombineOutputDialog.show(
                    frame,
                    inputDir,
                    rememberedOutputPath,
                    rememberedOutputFolder,
                    rememberedOpenWorkbook,
                    rememberedCopyCopilot);
            if (choice.isEmpty()) {
                return;
            }
            CombineOutputDialog.Result out = choice.get();
            rememberedOutputPath = out.explicitOutput() == null ? "" : out.explicitOutput().toString();
            rememberedOutputFolder = out.outputDirWhenAuto() == null ? "" : out.outputDirWhenAuto().toString();
            rememberedOpenWorkbook = out.openWorkbookAfter();
            rememberedCopyCopilot = out.copyCopilotPrompt();
            runCombine(frame, runButton, statusIndicator, inputDirField, recursiveBox, out, logArea);
        });

        frame.setSize(1080, 760);
        frame.setMinimumSize(new Dimension(900, 640));
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
        SwingUtilities.invokeLater(() -> split.setDividerLocation(0.58));
    }

    private static JPanel cardPanel(String title, JComponent body, JComponent headerTrailing) {
        JPanel outer = new JPanel(new BorderLayout(0, 12));
        outer.setOpaque(true);
        outer.setBackground(UiTheme.CARD_BG);
        outer.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createCompoundBorder(
                        BorderFactory.createMatteBorder(0, 4, 0, 0, UiTheme.CARD_ACCENT_STRIPE),
                        BorderFactory.createLineBorder(UiTheme.CARD_BORDER, 1, true)),
                BorderFactory.createEmptyBorder(CARD_PAD, CARD_PAD - 2, CARD_PAD + 2, CARD_PAD)));
        JLabel head = new JLabel(title);
        head.setFont(head.getFont().deriveFont(Font.BOLD, 13f));
        head.setForeground(UiTheme.TEXT_PRIMARY);
        if (headerTrailing == null) {
            outer.add(head, BorderLayout.NORTH);
        } else {
            JPanel header = new JPanel(new BorderLayout(12, 0));
            header.setOpaque(false);
            header.add(head, BorderLayout.WEST);
            header.add(headerTrailing, BorderLayout.EAST);
            outer.add(header, BorderLayout.NORTH);
        }
        outer.add(body, BorderLayout.CENTER);
        return outer;
    }

    private static JLabel sourceFieldLabel(String text) {
        JLabel l = new JLabel(text);
        l.setFont(l.getFont().deriveFont(Font.PLAIN, 13f));
        l.setForeground(UiTheme.TEXT_SECONDARY);
        l.setHorizontalAlignment(SwingConstants.LEFT);
        Dimension size = new Dimension(SOURCE_LABEL_WIDTH, l.getPreferredSize().height);
        l.setPreferredSize(size);
        l.setMinimumSize(size);
        return l;
    }

    private static boolean validateCombineSource(JFrame frame, JTextField inputDirField) {
        String in = inputDirField.getText().trim();
        if (in.isEmpty()) {
            JOptionPane.showMessageDialog(frame, "Choose the folder with your Excel files.", "MPR Combiner", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        Path inputDir = Paths.get(in).toAbsolutePath().normalize();
        if (!Files.isDirectory(inputDir)) {
            JOptionPane.showMessageDialog(frame, "That path is not a folder:\n" + inputDir, "MPR Combiner", JOptionPane.ERROR_MESSAGE);
            return false;
        }
        return true;
    }

    private static void runCombine(
            JFrame frame,
            JButton runButton,
            StatusIndicator statusIndicator,
            JTextField inputDirField,
            JCheckBox recursiveBox,
            CombineOutputDialog.Result output,
            JTextArea logArea
    ) {
        Path inputDir = Paths.get(inputDirField.getText().trim()).toAbsolutePath().normalize();
        Path explicitOutput = output.explicitOutput();
        Path outputDirWhenAuto = output.outputDirWhenAuto();
        boolean recursive = recursiveBox.isSelected();
        final boolean openWorkbookAfter = output.openWorkbookAfter();
        final boolean copyCopilotPrompt = output.copyCopilotPrompt();

        final Path[] written = new Path[1];

        logArea.setText("");
        statusIndicator.showWorking();
        runButton.setEnabled(false);

        SwingWorker<Void, Void> worker = new SwingWorker<>() {
            @Override
            protected Void doInBackground() throws Exception {
                int timeoutMin = combineTimeoutMinutes();
                ExecutorService pool = Executors.newSingleThreadExecutor(r -> {
                    Thread t = new Thread(r, "mpr-combine");
                    t.setDaemon(true);
                    return t;
                });
                Future<Path> future = pool.submit(() -> {
                    final Path[] out = new Path[1];
                    redirectLog(logArea, () -> {
                        out[0] = MPRCombiner.runLocalCombine(
                                inputDir,
                                explicitOutput,
                                outputDirWhenAuto,
                                DEFAULT_SHEET_NAME,
                                DEFAULT_HEADER_ROW,
                                recursive,
                                false,
                                false);
                    });
                    return out[0];
                });
                try {
                    written[0] = future.get(timeoutMin, TimeUnit.MINUTES);
                } catch (TimeoutException e) {
                    future.cancel(true);
                    SwingUtilities.invokeLater(() ->
                            logArea.append("\n[Timed out — combine stopped. See error dialog.]\n"));
                    throw new IllegalStateException(
                            "Timed out after " + timeoutMin + " minutes while reading or combining workbooks.\n\n"
                                    + "If these files are in OneDrive or SharePoint, wait until they are fully downloaded "
                                    + "on this computer (no “cloud only” placeholders), then try again.\n\n"
                                    + "To allow more time, start the app with e.g. "
                                    + "-Dmpr.combiner.timeoutMinutes=30");
                } catch (ExecutionException e) {
                    Throwable c = e.getCause();
                    if (c instanceof Exception ex) {
                        throw ex;
                    }
                    throw new RuntimeException(c);
                } catch (InterruptedException e) {
                    future.cancel(true);
                    Thread.currentThread().interrupt();
                    throw new IllegalStateException("Combine was interrupted.");
                } finally {
                    pool.shutdownNow();
                    try {
                        pool.awaitTermination(30, TimeUnit.SECONDS);
                    } catch (InterruptedException e) {
                        Thread.currentThread().interrupt();
                    }
                }
                return null;
            }

            @Override
            protected void done() {
                runButton.setEnabled(true);
                try {
                    get();
                    statusIndicator.showFinished();
                    Path out = written[0];
                    StringBuilder followUp = new StringBuilder();
                    if (out != null && copyCopilotPrompt) {
                        try {
                            String prompt = CopilotExcelPrompt.load();
                            Toolkit.getDefaultToolkit().getSystemClipboard().setContents(new StringSelection(prompt), null);
                            followUp.append("Copilot cleaning prompt copied to clipboard — paste into Copilot (⌘V / Ctrl+V).\n");
                        } catch (Exception clipEx) {
                            followUp.append("Could not copy Copilot prompt: ").append(clipEx.getMessage()).append('\n');
                        }
                    }
                    if (out != null && openWorkbookAfter && out.toString().toLowerCase(Locale.ROOT).endsWith(".xlsx")) {
                        if (Desktop.isDesktopSupported()) {
                            try {
                                Desktop.getDesktop().open(out.toFile());
                                followUp.append("Opened workbook with the default application for .xlsx files.\n");
                            } catch (Exception openEx) {
                                followUp.append("Could not open workbook: ").append(openEx.getMessage()).append('\n');
                            }
                        } else {
                            followUp.append("Opening files from the desktop is not supported on this system.\n");
                        }
                    }
                    String msg = out != null ? "Wrote combined file:\n" + out : "Done.";
                    if (followUp.length() > 0) {
                        msg = msg + "\n\n" + followUp.toString().trim();
                    }
                    JOptionPane.showMessageDialog(frame, msg, "Done", JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception ex) {
                    statusIndicator.showError();
                    JOptionPane.showMessageDialog(frame, unwrap(ex).getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                }
                statusIndicator.showReady();
            }
        };
        worker.execute();
    }

    private static void redirectLog(JTextArea logArea, ThrowingRunnable task) throws Exception {
        PrintStream oldOut = System.out;
        PrintStream oldErr = System.err;
        PrintStream gui = printStreamFor(logArea);
        System.setOut(gui);
        System.setErr(gui);
        try {
            task.run();
        } finally {
            System.setOut(oldOut);
            System.setErr(oldErr);
            gui.close();
        }
    }

    @FunctionalInterface
    private interface ThrowingRunnable {
        void run() throws Exception;
    }

    private static Throwable unwrap(Exception ex) {
        Throwable c = ex.getCause();
        return c != null ? c : ex;
    }

    private static PrintStream printStreamFor(JTextArea area) {
        OutputStream os = new OutputStream() {
            @Override
            public void write(int b) {
                write(new byte[]{(byte) b}, 0, 1);
            }

            @Override
            public void write(byte[] b, int off, int len) {
                if (len <= 0) {
                    return;
                }
                String chunk = new String(b, off, len, StandardCharsets.UTF_8);
                SwingUtilities.invokeLater(() -> {
                    area.append(chunk);
                    area.setCaretPosition(area.getDocument().getLength());
                });
            }
        };
        return new PrintStream(os, true, StandardCharsets.UTF_8);
    }

    private static boolean isExcelFileName(String name) {
        int dot = name.lastIndexOf('.');
        if (dot < 0) {
            return false;
        }
        String ext = name.substring(dot + 1).toLowerCase(Locale.ROOT);
        return EXCEL_EXT.contains(ext);
    }

    private static Comparator<Path> syncedTreePathOrder() {
        return (a, b) -> {
            boolean da = Files.isDirectory(a);
            boolean db = Files.isDirectory(b);
            if (da != db) {
                return da ? -1 : 1;
            }
            String na = a.getFileName() != null ? a.getFileName().toString() : a.toString();
            String nb = b.getFileName() != null ? b.getFileName().toString() : b.toString();
            if (!da) {
                boolean ea = isExcelFileName(na);
                boolean eb = isExcelFileName(nb);
                if (ea != eb) {
                    return ea ? -1 : 1;
                }
            }
            return na.compareToIgnoreCase(nb);
        };
    }

    private static void pickSyncedLibraryRoot(
            JFrame frame,
            JTree tree,
            DefaultTreeModel model,
            JTextField inputDirField,
            JPanel treePanel) {
        JFileChooser chooser = new JFileChooser();
        configureFileChooser(chooser);
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.setDialogTitle("Select synced library (e.g. OneDrive or SharePoint folder)");
        Path home = Paths.get(System.getProperty("user.home", "."));
        Path[] quickStarts = {
                home.resolve("Library/CloudStorage"),
                home.resolve("OneDrive"),
                home
        };
        for (Path p : quickStarts) {
            if (Files.isDirectory(p)) {
                chooser.setCurrentDirectory(p.toFile());
                break;
            }
        }
        if (chooser.showOpenDialog(frame) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        Path root = chooser.getSelectedFile().toPath().toAbsolutePath().normalize();
        if (!Files.isDirectory(root)) {
            JOptionPane.showMessageDialog(frame, "Not a folder:\n" + root, "MPR Combiner", JOptionPane.ERROR_MESSAGE);
            return;
        }
        loadLibraryTree(tree, model, inputDirField, root, treePanel);
    }

    private static void loadLibraryTree(
            JTree tree,
            DefaultTreeModel model,
            JTextField inputDirField,
            Path root,
            JPanel treePanel) {
        DefaultMutableTreeNode rootNode = new DefaultMutableTreeNode(new FileTreePayload(root, true));
        rootNode.add(new DefaultMutableTreeNode(TREE_PLACEHOLDER));
        model.setRoot(rootNode);
        tree.setRootVisible(true);
        tree.setShowsRootHandles(true);
        ((CardLayout) treePanel.getLayout()).show(treePanel, "tree");
        inputDirField.setText(root.toString());
        SwingUtilities.invokeLater(() -> {
            tree.expandRow(0);
            tree.setSelectionRow(0);
        });
    }

    static boolean pickDirectory(Component parent, JTextField target, Path optionalStart, String dialogTitle) {
        JFileChooser chooser = new JFileChooser();
        configureFileChooser(chooser);
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.setDialogTitle(dialogTitle);
        if (optionalStart != null && Files.isDirectory(optionalStart)) {
            chooser.setCurrentDirectory(optionalStart.toFile());
        }
        if (chooser.showOpenDialog(parent) != JFileChooser.APPROVE_OPTION) {
            return false;
        }
        target.setText(chooser.getSelectedFile().getAbsolutePath());
        return true;
    }

    static void pickOutputFile(Component owner, JTextField outputField, Path inputDirForDefault) {
        JFileChooser chooser = new JFileChooser();
        configureFileChooser(chooser);
        chooser.setDialogTitle("Save combined workbook");
        chooser.setAcceptAllFileFilterUsed(true);
        chooser.addChoosableFileFilter(new FileNameExtensionFilter("Excel (*.xlsx)", "xlsx"));
        chooser.addChoosableFileFilter(new FileNameExtensionFilter("CSV (*.csv)", "csv"));
        if (inputDirForDefault != null && Files.isDirectory(inputDirForDefault)) {
            chooser.setCurrentDirectory(inputDirForDefault.toFile());
        }
        String cur = outputField.getText().trim();
        if (!cur.isEmpty()) {
            Path p = Paths.get(cur);
            Path parentDir = p.getParent();
            if (parentDir != null && Files.isDirectory(parentDir)) {
                chooser.setCurrentDirectory(parentDir.toFile());
            }
            if (p.getFileName() != null) {
                chooser.setSelectedFile(p.toFile());
            }
        }
        if (chooser.showSaveDialog(owner) == JFileChooser.APPROVE_OPTION) {
            outputField.setText(chooser.getSelectedFile().getAbsolutePath());
        }
    }

    private static final class FileTreePayload {
        final Path path;
        final boolean directory;
        volatile boolean loading;

        FileTreePayload(Path path, boolean directory) {
            this.path = path;
            this.directory = directory;
        }

        @Override
        public String toString() {
            String name = path.getFileName() != null ? path.getFileName().toString() : path.toString();
            return directory ? name + " /" : name;
        }
    }

    private static final class LocalFolderTreeExpansionListener implements TreeExpansionListener {
        private final JTree tree;
        private final DefaultTreeModel model;

        LocalFolderTreeExpansionListener(JTree tree, DefaultTreeModel model) {
            this.tree = tree;
            this.model = model;
        }

        private static boolean onlyPlaceholderChild(DefaultMutableTreeNode node) {
            return node.getChildCount() == 1
                    && ((DefaultMutableTreeNode) node.getFirstChild()).getUserObject() == TREE_PLACEHOLDER;
        }

        @Override
        public void treeExpanded(TreeExpansionEvent event) {
            DefaultMutableTreeNode node = (DefaultMutableTreeNode) event.getPath().getLastPathComponent();
            Object u = node.getUserObject();
            if (!(u instanceof FileTreePayload p) || !p.directory) {
                return;
            }
            if (p.loading || !onlyPlaceholderChild(node)) {
                return;
            }
            p.loading = true;
            Path dir = p.path;

            SwingWorker<List<Path>, Void> worker = new SwingWorker<>() {
                @Override
                protected List<Path> doInBackground() throws IOException {
                    try (Stream<Path> stream = Files.list(dir)) {
                        return stream
                                .filter(p -> !isHiddenPath(p))
                                .sorted(syncedTreePathOrder())
                                .toList();
                    }
                }

                @Override
                protected void done() {
                    try {
                        List<Path> kids = get();
                        node.removeAllChildren();
                        for (Path child : kids) {
                            boolean isDir = Files.isDirectory(child);
                            DefaultMutableTreeNode cn = new DefaultMutableTreeNode(new FileTreePayload(child, isDir));
                            if (isDir) {
                                cn.add(new DefaultMutableTreeNode(TREE_PLACEHOLDER));
                            }
                            node.add(cn);
                        }
                        model.reload(node);
                    } catch (Exception ex) {
                        node.removeAllChildren();
                        node.add(new DefaultMutableTreeNode(TREE_PLACEHOLDER));
                        model.reload(node);
                        Throwable t = unwrap(ex);
                        SwingUtilities.invokeLater(() ->
                                JOptionPane.showMessageDialog(tree, t.getMessage(), "Could not read folder", JOptionPane.ERROR_MESSAGE));
                    } finally {
                        p.loading = false;
                    }
                }
            };
            worker.execute();
        }

        @Override
        public void treeCollapsed(TreeExpansionEvent event) {
        }
    }

    private static final class StatusIndicator extends JPanel {
        private enum Mode {
            READY, WORKING, FINISHED, ERROR
        }

        private static final int GLYPH_SIZE = 20;
        private static final int DOT_SIZE = 8;
        private static final float WORKING_STROKE = 2f;

        private Mode mode = Mode.READY;
        private int spinAngle;
        private float readyPhase;
        private final JLabel message = new JLabel("Ready");
        private final JPanel glyph = new JPanel() {
            @Override
            protected void paintComponent(Graphics g) {
                Graphics2D g2 = (Graphics2D) g.create();
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                int w = getWidth();
                int h = getHeight();
                int cx = w / 2;
                int cy = h / 2;
                switch (mode) {
                    case WORKING -> paintWorkingGlyph(g2, cx, cy, Math.min(w, h));
                    case ERROR -> paintStatusDot(g2, cx, cy, UiTheme.RED);
                    case FINISHED -> paintStatusDot(g2, cx, cy, UiTheme.STATUS_DONE);
                    default -> paintReadyGlyph(g2, cx, cy);
                }
                g2.dispose();
            }
        };
        private final Timer animation = new Timer(36, e -> {
            if (mode == Mode.READY) {
                readyPhase += 0.02f;
                if (readyPhase >= 1f) {
                    readyPhase -= 1f;
                }
                glyph.repaint();
            } else if (mode == Mode.WORKING) {
                spinAngle = (spinAngle + 26) % 360;
                glyph.repaint();
            }
        });

        StatusIndicator() {
            setOpaque(false);
            setLayout(new GridBagLayout());
            JPanel row = new JPanel();
            row.setOpaque(false);
            row.setLayout(new BoxLayout(row, BoxLayout.LINE_AXIS));
            glyph.setOpaque(false);
            glyph.setPreferredSize(new Dimension(GLYPH_SIZE, GLYPH_SIZE));
            glyph.setAlignmentY(Component.CENTER_ALIGNMENT);
            message.setFont(message.getFont().deriveFont(Font.PLAIN, 13f));
            message.setForeground(UiTheme.TEXT_MUTED);
            message.setVerticalAlignment(SwingConstants.CENTER);
            message.setAlignmentY(Component.CENTER_ALIGNMENT);
            row.add(glyph);
            row.add(Box.createHorizontalStrut(8));
            row.add(message);
            GridBagConstraints statusGc = new GridBagConstraints();
            statusGc.gridx = 0;
            statusGc.gridy = 0;
            statusGc.anchor = GridBagConstraints.CENTER;
            add(row, statusGc);
            int actionHeight = UiTheme.actionButtonHeight();
            Dimension size = new Dimension(row.getPreferredSize().width, actionHeight);
            setPreferredSize(size);
            setMinimumSize(size);
            setMaximumSize(new Dimension(Integer.MAX_VALUE, actionHeight));
            animation.setRepeats(true);
            animation.setCoalesce(true);
            animation.start();
        }

        void showReady() {
            mode = Mode.READY;
            message.setText("Ready");
            message.setForeground(UiTheme.TEXT_MUTED);
            if (!animation.isRunning()) {
                animation.start();
            }
            glyph.repaint();
        }

        void showWorking() {
            mode = Mode.WORKING;
            spinAngle = 0;
            message.setText("Working…");
            message.setForeground(UiTheme.TEXT_SECONDARY);
            if (!animation.isRunning()) {
                animation.start();
            }
            glyph.repaint();
        }

        void showFinished() {
            mode = Mode.FINISHED;
            animation.stop();
            message.setText("Finished");
            message.setForeground(UiTheme.STATUS_DONE);
            glyph.repaint();
        }

        void showError() {
            mode = Mode.ERROR;
            animation.stop();
            message.setText("Error");
            message.setForeground(UiTheme.RED);
            glyph.repaint();
        }

        private void paintWorkingGlyph(Graphics2D g2, int cx, int cy, int extent) {
            g2.setStroke(new BasicStroke(WORKING_STROKE, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
            g2.setColor(UiTheme.RED);
            int arc = Math.max(6, extent - (int) Math.ceil(WORKING_STROKE) - 2);
            int x = cx - arc / 2;
            int y = cy - arc / 2;
            g2.drawArc(x, y, arc, arc, spinAngle, 270);
        }

        private void paintStatusDot(Graphics2D g2, int cx, int cy, Color color) {
            g2.setColor(color);
            int r = DOT_SIZE / 2;
            g2.fillOval(cx - r, cy - r, DOT_SIZE, DOT_SIZE);
        }

        private void paintReadyGlyph(Graphics2D g2, int cx, int cy) {
            Color ready = UiTheme.STATUS_READY;
            double breathe = 0.5 + 0.5 * Math.sin(readyPhase * Math.PI * 2);
            int dot = (int) (DOT_SIZE - 1 + breathe * 2);
            float alpha = (float) (0.72 + 0.28 * breathe);
            g2.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER, alpha));
            g2.setColor(ready);
            g2.fillOval(cx - dot / 2, cy - dot / 2, dot, dot);
            g2.setComposite(AlphaComposite.SrcOver);
        }
    }

    private static final class BrandHeader extends JPanel {
        BrandHeader(Optional<ImageIcon> brandMark) {
            setLayout(new BorderLayout());
            setOpaque(false);
            int padY = 20;
            int padX = 28;

            JLabel title = new JLabel("MPR Combiner");
            title.setForeground(Color.WHITE);
            title.setFont(title.getFont().deriveFont(Font.BOLD, 25f));

            JPanel row = new JPanel();
            row.setOpaque(false);
            row.setLayout(new BoxLayout(row, BoxLayout.X_AXIS));
            row.setBorder(BorderFactory.createEmptyBorder(padY, padX, padY, padX));
            brandMark.ifPresent(icon -> {
                JLabel logo = new JLabel(icon);
                logo.setOpaque(false);
                logo.setAlignmentY(Component.TOP_ALIGNMENT);
                row.add(logo);
                row.add(Box.createHorizontalStrut(22));
            });
            row.add(title);
            add(row, BorderLayout.WEST);
            setMinimumSize(new Dimension(10, brandMark.isPresent() ? 76 : 68));
        }

        @Override
        protected void paintComponent(Graphics g) {
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            int w = getWidth();
            int h = getHeight();
            g2.setPaint(new GradientPaint(0, 0, UiTheme.HEADER_TOP, w, h, UiTheme.HEADER_BOTTOM));
            g2.fillRect(0, 0, w, h);
            g2.setColor(new Color(255, 255, 255, 12));
            g2.fillOval(w - 200, -90, 360, 240);
            g2.setColor(new Color(255, 255, 255, 18));
            g2.fillOval(-100, h - 36, 260, 120);
            g2.setColor(new Color(0, 0, 0, 35));
            g2.drawLine(0, h - 1, w, h - 1);
            g2.dispose();
            super.paintComponent(g);
        }
    }
}
