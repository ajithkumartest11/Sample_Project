package com.selenium.crm.pages;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;

import com.selenium.crm.baseclass.TestBase;

public class LoginPage extends TestBase{

	@FindBy(name="username")
	WebElement username;
	
	@FindBy(name="password")
	WebElement password;
	
	@FindBy(xpath="//input[@class='btn btn-small' and @type='submit']") 
	WebElement loginButton;
	
	//@FindBy(xpath="//button[contains(text(),'Sign Up')]")
	//WebElement signUpButton;
	
	@FindBy(linkText = "Sign Up")
	WebElement signUpButton;
	
	//@FindBy(xpath="//img[contains(@class,'img-responsive')]")
	@FindBy(xpath = "(//img[contains(@class,'img-responsive')])[1]")
	WebElement crmLogo;	
	
		
	public LoginPage()
	{
		PageFactory.initElements(driver, this);
	}
	
	public String validateLoginPageTitle()
	{
		return driver.getTitle();
	}
	
	public boolean validateCRMImage()
	{
		return crmLogo.isDisplayed();
	}
	
	public HomePage login(String uname, String pword)
	{
		username.sendKeys(uname);
		password.sendKeys(pword);
		
		JavascriptExecutor js = (JavascriptExecutor)driver;
		js.executeScript("arguments[0].click();", loginButton);
		
				
		return new HomePage();
	}
	
}
