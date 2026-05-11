package com.britan.mprcombiner;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertTrue;

class CopilotExcelPromptTest {

    @Test
    void bundledPromptContainsKeyInstructions() throws Exception {
        String s = CopilotExcelPrompt.load();
        assertTrue(s.contains("Processor ID"));
        assertTrue(s.contains("Used LBS"));
    }
}
