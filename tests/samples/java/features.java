/** A point in two-dimensional space. */
public record Point(double x, double y) {}

class Features {
    private long total$ = 1_000_000L;
    private int flags = 0b1010; // Selects the switch branch.
    private String json = """
            {"enabled": true}
            """;

    /* Blocks can span
       more than one line. */
    @SuppressWarnings("unchecked") List<String> values() {
        // Calls, generics, annotations, and switch expressions.
        return switch (flags) {
            case 0 -> List.of();
            default -> List.of(String.valueOf(total$));
        };
    }
}
