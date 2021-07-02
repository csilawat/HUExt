package opps;

public class PalindromeNumber {

    public static void main(String[] args) {

        String inputStringFirst = "saas";
        String inputStringSecond = "ababc";
        String outputString = "";

        int length = inputStringFirst.length();
        for (int i = length - 1; i >= 0; i--) {
            outputString = outputString + inputStringFirst.charAt(i);
        }

        if (inputStringFirst.equals(outputString))
            System.out.println(inputStringFirst + " is a palindrome");
        else {
            System.out.println(inputStringFirst + " is not a palindrome");
        }

        int swap = countSwap(inputStringSecond);
        if (swap == -1) {
            System.out.println("Given String can not be converted into palindrome");
        } else {
            System.out.println(swap + " swap required to converted '"+inputStringSecond+"' into palindrome ");
        }
    }


    static int countSwap(String str) {
        int n = str.length();
        char s[] = str.toCharArray();
        int count = 0;

        for (int i = 0; i < n / 2; i++) {

            int left = i;
            int right = n - left - 1;
            while (left < right) {
                if (s[left] == s[right]) {
                    break;
                } else {
                    right--;
                }
            }
            if (left == right) {
                return -1;
            } else if (s[left] != s[n - left - 1]) {
                char temp = s[right];
                s[right] = s[n - left - 1];
                s[n - left - 1] = temp;
                count++;
            }
        }

        return count;
    }

}
