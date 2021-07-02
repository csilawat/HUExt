package collections;

import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

public class SortHashMap {

    public static void main(String[] args) {

        Map<String, Integer> nameMap = new HashMap<String, Integer>();
        nameMap.put("Jayant", 80);
        nameMap.put("Anushka", 80);
        nameMap.put("Amit", 75);
        nameMap.put("Abhishek", 90);
        nameMap.put("Danish", 40);

        shortedMap(nameMap);

    }

    public static void shortedMap(Map<String, Integer> nameMap) {

        TreeMap<String, Integer> shortedMap = new TreeMap<String, Integer>(nameMap);

        for (Map.Entry<String, Integer> entry : shortedMap.entrySet())
            System.out.println("Key = " + entry.getKey() +
                    ", Value = " + entry.getValue());
    }

}
