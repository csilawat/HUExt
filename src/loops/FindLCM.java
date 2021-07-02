package loops;

public class FindLCM {

    public static void main(String[] args) {

        int lcm = findLCM(15, 50);
        System.out.println("LCM of given numbers is :: " + lcm);
    }

    public static int findLCM(int firstNumber, int secondNumber) {

        int max, maxPlus, lcm = 0;

        if (firstNumber > secondNumber) {
            max = maxPlus = firstNumber;
        } else {
            max = maxPlus = secondNumber;
        }

        while (firstNumber != 0) {
            if (max % firstNumber == 0 && max % secondNumber == 0) {
                lcm = max;
                break;
            }
            max += maxPlus;
        }
        return lcm;
    }

}
