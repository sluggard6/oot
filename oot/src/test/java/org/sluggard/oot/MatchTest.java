package org.sluggard.oot;

import java.util.regex.Pattern;

import org.junit.jupiter.api.Test;

public class MatchTest {
	
	@Test
	public void hasNumber() {
		String[] testString = {"aaa1bbb", "1abcd", "acbd1" , "adbcde", "abce_daba", "abce-daba"};
		Pattern pattern = Pattern.compile("^[A-Za-z_]+$");
		for(String s : testString) {
			System.out.println(pattern.matcher(s).matches());
		}
	}

}
