package demo;

import java.util.ArrayList;
import java.util.List;

public final class Greeter {
    private final List<String> names = new ArrayList<>();

    @Deprecated
    public void add(String name) {
        if (name == null) {
            throw new IllegalArgumentException("name");
        }
        names.add(name);
    }

    public String greeting(int index) {
        return "Hello, " + names.get(index) + "!";
    }
}
