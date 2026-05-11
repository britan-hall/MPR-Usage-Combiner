package com.britan.mprcombiner;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

/** Text bundled for Excel Copilot (cleaning the combined Usage sheet). */
final class CopilotExcelPrompt {

    private static volatile String cached;

    private CopilotExcelPrompt() {
    }

    static String load() throws IOException {
        if (cached != null) {
            return cached;
        }
        try (InputStream in = CopilotExcelPrompt.class.getResourceAsStream("/copilot-excel-prompt.txt")) {
            if (in == null) {
                throw new IOException("Missing classpath resource: copilot-excel-prompt.txt");
            }
            cached = new String(in.readAllBytes(), StandardCharsets.UTF_8).trim();
            return cached;
        }
    }
}
