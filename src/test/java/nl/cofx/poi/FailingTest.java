package nl.cofx.poi;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertFalse;

public class FailingTest {

    @Test
    void failImmediately() {
        assertFalse(true);
    }
}
