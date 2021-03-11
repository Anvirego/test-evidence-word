package com.github.anvirego;

import com.github.anvirego.interfaces.TestEvidenceInterface;

/**
 * @author Ing. Angélica Viridiana Rebolloza Gonzalez.
 * @version 1.0 03/2021.
 * TestEvidenceInstance: Defines and creates test evidence's instances to work with the Interface implemented.
 */
public class TestEvidenceInstance {
	public static TestEvidenceInterface getInstance(String platform, String suiteTestName, String testerName, String environment) {
		TestEvidenceInterface ti = TestEvidence.getInstance(platform, suiteTestName, testerName, environment);
		return ti;
	}//Method

}//Class