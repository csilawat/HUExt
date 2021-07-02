package loops;

public class FindFrequencyOfInteger {

    public static void main(String[] args) {
        int inputNumber = 32143214;
        findFrequency(inputNumber);
    }

    public static void findFrequency(int inputNumber) {

        int[] digitCount = new int[10];

        while (inputNumber != 0) {
            int lastDigit = inputNumber % 10;
            digitCount[lastDigit]++;
            inputNumber = inputNumber / 10;
        }
        for (int i = 0; i < digitCount.length; i++) {
            if (digitCount[i] != 0) {
                System.out.println(i + " - " + digitCount[i]);
            }
        }
    }
}
