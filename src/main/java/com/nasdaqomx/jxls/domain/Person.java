package com.nasdaqomx.jxls.domain;

/**
 *
 * @author Ola Theander <ola.theander@nasdaqomx.com>
 */
public class Person {
    private String firstName;
    private String surName;
    private Integer age;

    public Person() {
    }
    
    public Person(String firstName, String surName, Integer age) {
        this.firstName = firstName;
        this.surName = surName;
        this.age = age;
    }

    public String getFirstName() {
        return firstName;
    }

    public void setFirstName(String firstName) {
        this.firstName = firstName;
    }

    public String getSurName() {
        return surName;
    }

    public void setSurName(String surName) {
        this.surName = surName;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }
}
