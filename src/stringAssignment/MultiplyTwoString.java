package stringAssignment;

public class MultiplyTwoString {
    public static void main(String[] args) {

        String firstValue = "12";
        String secondValue = "2";

        int firstIntValue = Integer.parseInt(secondValue);
        int secondIntValue = Integer.parseInt(firstValue);

        int multiplyValue = firstIntValue * secondIntValue;
        String stringValue = String.valueOf(multiplyValue);

        System.out.println(stringValue);
    }
}
