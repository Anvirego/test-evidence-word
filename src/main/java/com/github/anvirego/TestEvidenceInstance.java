package com.github.anvirego;

import com.github.anvirego.interfaces.TestEvidenceInterface;

public class TestEvidenceInstance {
	public static TestEvidenceInterface getInstance(String platform, String suiteTestName, String testerName, String environment) {
		TestEvidenceInterface ti = TestEvidence.getInstance(platform, suiteTestName, testerName, environment);
		return ti;
	}//Method

}//Class
