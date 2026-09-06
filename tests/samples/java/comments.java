// Line comments can contain keywords, "strings", and /* block markers */.
/**
 * Javadoc is highlighted as a comment, including @param and {@code true}.
 */
class Comments {
    /* A single-line block comment. */
    int count = 1; // A trailing line comment.
    /* A multiline block comment with // inside it.
       Keywords like class and return remain comment text.
     */
    int next = /* An inline comment. */ 2;
    /**/ int empty = 0;
    String url = "https://example.com";
    String markers = "/* This is a string, not a comment. */";
    String text = """
            // These markers are inside a text block.
            /* They remain string text. */
            """;
    int quotient = 8 / 2;
} // Comments also follow closing braces.
