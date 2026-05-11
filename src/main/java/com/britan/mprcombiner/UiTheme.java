package com.britan.mprcombiner;

import com.formdev.flatlaf.FlatClientProperties;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import java.awt.Color;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GradientPaint;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.RenderingHints;

/** Shared red / gray / white palette and controls for the desktop UI. */
final class UiTheme {
    static final Color BG_APP = new Color(0xee_ee_ee);
    static final Color CARD_BG = Color.WHITE;
    static final Color CARD_BORDER = new Color(0xd4_d4_d4);
    static final Color CARD_ACCENT_STRIPE = new Color(0xb9_1c_1c);

    static final Color LOG_SURFACE = new Color(0xfa_fa_fa);
    static final Color LOG_BORDER = new Color(0xe5_e5_e5);

    static final Color TEXT_PRIMARY = new Color(0x17_17_17);
    static final Color TEXT_SECONDARY = new Color(0x52_52_52);
    static final Color TEXT_MUTED = new Color(0x73_73_73);

    static final Color RED = new Color(0xb9_1c_1c);

    static final Color HEADER_TOP = new Color(0x26_26_26);
    static final Color HEADER_BOTTOM = new Color(0x7f_1d_1d);

    static final Color STATUS_DONE = new Color(0x40_40_40);
    static final Color STATUS_READY = new Color(0x16_a34a);

    private static final int BUTTON_HEIGHT = 38;
    private static final int COMPACT_BUTTON_HEIGHT = 32;
    private static final int BUTTON_PAD_X = 18;
    private static final int COMPACT_BUTTON_PAD_X = 10;
    private static final int BUTTON_CORNER_ARC = 8;

    private UiTheme() {
    }

    /** Same charcoal-to-red gradient as the app header bar. */
    static JButton primaryButton(String text) {
        JButton b = new JButton(text) {
            @Override
            protected void paintComponent(Graphics g) {
                Graphics2D g2 = (Graphics2D) g.create();
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                int w = getWidth();
                int h = getHeight();
                if (!isEnabled()) {
                    g2.setColor(new Color(0x99_99_99));
                    g2.fillRoundRect(0, 0, w, h, BUTTON_CORNER_ARC, BUTTON_CORNER_ARC);
                } else {
                    Color top = HEADER_TOP;
                    Color bottom = HEADER_BOTTOM;
                    if (getModel().isPressed()) {
                        top = HEADER_BOTTOM;
                        bottom = HEADER_TOP;
                    } else if (getModel().isRollover()) {
                        top = HEADER_TOP.brighter();
                        bottom = HEADER_BOTTOM.brighter();
                    }
                    g2.setPaint(new GradientPaint(0, 0, top, w, h, bottom));
                    g2.fillRoundRect(0, 0, w, h, BUTTON_CORNER_ARC, BUTTON_CORNER_ARC);
                }
                g2.dispose();
                super.paintComponent(g);
            }
        };
        b.setForeground(Color.WHITE);
        b.setContentAreaFilled(false);
        b.setBorderPainted(false);
        b.setFocusPainted(false);
        b.setFont(b.getFont().deriveFont(Font.BOLD, 13f));
        b.putClientProperty(FlatClientProperties.BUTTON_TYPE, FlatClientProperties.BUTTON_TYPE_ROUND_RECT);
        b.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        b.setMinimumSize(new Dimension(96, BUTTON_HEIGHT));
        b.setPreferredSize(new Dimension(Math.max(120, b.getPreferredSize().width + BUTTON_PAD_X), BUTTON_HEIGHT));
        return b;
    }

    static JButton secondaryButton(String text) {
        return outlineButton(text, BUTTON_HEIGHT, BUTTON_PAD_X, 88, 100, false);
    }

    static JButton compactSecondaryButton(String text) {
        return outlineButton(text, COMPACT_BUTTON_HEIGHT, COMPACT_BUTTON_PAD_X, 0, 0, true);
    }

    static int actionButtonHeight() {
        return BUTTON_HEIGHT;
    }

    private static JButton outlineButton(
            String text,
            int height,
            int padX,
            int minWidth,
            int preferredMinWidth,
            boolean lockWidth) {
        JButton b = new JButton(text);
        b.setForeground(TEXT_SECONDARY);
        b.setBackground(CARD_BG);
        b.setOpaque(true);
        b.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(CARD_BORDER, 1, true),
                BorderFactory.createEmptyBorder(0, padX, 0, padX)));
        b.setFocusPainted(false);
        b.setFont(b.getFont().deriveFont(Font.PLAIN, 13f));
        b.putClientProperty(FlatClientProperties.BUTTON_TYPE, FlatClientProperties.BUTTON_TYPE_ROUND_RECT);
        b.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        int width = b.getPreferredSize().width;
        int prefWidth = preferredMinWidth > 0 ? Math.max(preferredMinWidth, width) : width;
        Dimension pref = new Dimension(prefWidth, height);
        b.setPreferredSize(pref);
        b.setMinimumSize(new Dimension(minWidth > 0 ? Math.max(minWidth, prefWidth) : prefWidth, height));
        if (lockWidth) {
            b.setMaximumSize(pref);
        }
        return b;
    }
}
