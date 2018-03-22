package br.com.inmetrics.RegistroPonto.Apoio;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

/**
 * Classe responsável por definir configuração do navegador
 * @author Ricardo Kohatsu
 *
 */
public class Navegador {
	private String nomeNavegador;
	
	public String getNomeNavegador() {
		return nomeNavegador;
	}

	public void setNomeNavegador(String nomeNavegador) {
		this.nomeNavegador = nomeNavegador;
	}

	public WebDriver configurarNavegador() {
	     switch (nomeNavegador.toUpperCase()) {  
	       case "CHROME":
	    	   ChromeOptions options= new ChromeOptions();
	    	   options.addArguments("--disable-extensions");
		   		System.setProperty("webdriver.chrome.driver", "drivers/chromedriver.exe");
				return new ChromeDriver(options);
	       default: 
	    	   System.out.println("Tipo de navegador não conhecido");  
	       	   return null;
	     }
	}
	
	public void acessarPaginaWeb(String url, WebDriver driver) {
		driver.get(url);
	}
}
