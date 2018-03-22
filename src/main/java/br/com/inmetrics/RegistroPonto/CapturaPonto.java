package br.com.inmetrics.RegistroPonto;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Before;
import org.junit.Test;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

import br.com.inmetrics.RegistroPonto.Apoio.ArquivoExcel;
import br.com.inmetrics.RegistroPonto.Apoio.Horario;
import br.com.inmetrics.RegistroPonto.Apoio.Navegador;
import br.com.inmetrics.RegistroPonto.Apoio.Ponto;
import br.com.inmetrics.RegistroPonto.Funcionalidades.FuncExcel;
import br.com.inmetrics.RegistroPonto.Funcionalidades.FuncZeus;
import junit.framework.TestCase;

 

public class CapturaPonto extends TestCase {

       WebDriver driver;
       Navegador navegador;

       @Before
       public void before() {
       
       }
       
       @Test
       public void test() throws IOException {
    	   navegador = new Navegador();
           navegador.setNomeNavegador("Chrome");
           driver = navegador.configurarNavegador();
           driver.manage().window().maximize();    
    	   
    	   FuncExcel excel = new FuncExcel();
    	   FuncZeus zeus = new FuncZeus(driver);
    	   
    	   	String nomeAnalista = "Ricardo Augusto de Arruda";
    	   	String nomeAnalista2 = "Alex Novais Freitas";
    	   	String nomeAnalista3 = "Simone Andrade da Silva";
    	   	List<String> nomeAnalistas = new ArrayList<String>();
    	   	nomeAnalistas.add(nomeAnalista);
    	   	//nomeAnalistas.add(nomeAnalista2);
    	   	//nomeAnalistas.add(nomeAnalista3);
    	   	File arquivo = new File("C:\\\\Users\\\\Public\\\\RegistroPonto\\\\PontoEquipe.xlsx");
    	   	String url = "https://aplic.inmetrics.com.br//requerimento/frequencia.php";
    	   	String usuario = "16097331851";
    	   	String senha = "16097331851";
    	   	String periodo = "02/2018";
 
    	   	navegador.acessarPaginaWeb(url, driver);

    	   	zeus.realizarLogin(usuario, senha);
    	   	
    	   	zeus.selecionarPeriodo(periodo);

    	   	excel.validarExcel(arquivo);
    	   	
    	   	for (String nome : nomeAnalistas) {
    	   		zeus.selecionarFuncionario(nome);
    	   		excel.gravarDadosNoExcel(zeus.capturarRegistrosPonto(),nome);
    	   	}
    	   	excel.salvarFinalizarExcel(arquivo);
       }

}