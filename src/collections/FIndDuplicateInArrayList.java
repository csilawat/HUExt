package collections;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class FIndDuplicateInArrayList {

    public static void main(String[] args) {
        List<List<Integer>> mainList = new ArrayList<>();
        List<Integer> firstList = new ArrayList<>();
        firstList.add(1);
        firstList.add(1);
        firstList.add(3);
        firstList.add(2);
        mainList.add(firstList);

        List<Integer> secondList = new ArrayList<>();
        secondList.add(9);
        secondList.add(8);
        secondList.add(8);
        secondList.add(1);
        mainList.add(secondList);

        List<Integer> thirdList = new ArrayList<>();
        thirdList.add(0);
        thirdList.add(4);
        thirdList.add(5);
        thirdList.add(0);
        thirdList.add(0);
        thirdList.add(1);
        thirdList.add(4);
        mainList.add(thirdList);

        findDuplicateInList(mainList);
    }

    public static void findDuplicateInList(List<List<Integer>> mainList) {

        List<Map<Integer, Integer>> hashMapList = new ArrayList<>();

        int count = 0;
        for (int i = 0; i < mainList.size(); i++) {
            Map<Integer, Integer> map = new HashMap<>();
            for (Integer arrayList : mainList.get(i)) {

                if (map.containsKey(arrayList)) {
                    count = map.get(arrayList.intValue());
                    count++;
                    map.put(arrayList, count);
                } else {
                    map.put(arrayList, 1);
                }
            }
            hashMapList.add(map);
        }

        for (int i = 0; i < hashMapList.size(); i++) {
            Map<Integer, Integer> map = hashMapList.get(i);

            for (Integer key : map.keySet()) {

                if (map.get(key) > 1) {
                    System.out.println(key + " -> " + map.get(key));
                }
            }
        }
    }

}
