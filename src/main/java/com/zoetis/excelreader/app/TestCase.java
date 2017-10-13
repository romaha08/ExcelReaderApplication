package com.zoetis.excelreader.app;

import lombok.Getter;
import lombok.Setter;
import java.io.Serializable;
import java.util.Objects;

public class TestCase implements Serializable {
    @Getter @Setter private String testName;
    @Getter @Setter private String description;
    @Getter @Setter private String requirement;
    @Getter @Setter private String testResult;

    @Override
    public String toString() {
        return "TestCase{" +
                "testName='" + testName + '\'' +
                ", description='" + description + '\'' +
                ", requirement='" + requirement + '\'' +
                ", testResult='" + testResult + '\'' +
                '}';
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        TestCase testCase = (TestCase) o;
        return Objects.equals(testName, testCase.testName) &&
                Objects.equals(description, testCase.description) &&
                Objects.equals(requirement, testCase.requirement) &&
                Objects.equals(testResult, testCase.testResult);
    }

    @Override
    public int hashCode() {
        return Objects.hash(testName, description, requirement, testResult);
    }
}
