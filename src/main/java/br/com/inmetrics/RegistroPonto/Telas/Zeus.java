package br.com.inmetrics.RegistroPonto.Telas;

import java.util.ArrayList;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

import br.com.inmetrics.RegistroPonto.Apoio.ArquivoExcel;
import br.com.inmetrics.RegistroPonto.Apoio.Horario;
import br.com.inmetrics.RegistroPonto.Apoio.Ponto;

public class Zeus {
	private WebDriver driver;
	public Zeus(WebDriver driver) {
		this.driver = driver;
	}
	
	//************************Tela de Login***********************
	public void preencherUsuario(String texto) {
		driver.findElement(By.id("fun_Id")).sendKeys(texto);
	}
	
	public void preencherSenha(String texto) {
		driver.findElement(By.id("fun_Senha")).sendKeys(texto);
	}	
	
	public void clicarBotaoContinuar() {
		driver.findElement(By.id("btnSubmitLogn")).click();
	}		
	
	//************************Tela Principal***********************
	
	public void clicarBotaoFrequencia() {
		driver.findElement(By.xpath("//a[@title='Frequencia']")).click();
	}		
	
	public void clicarBotaoAlterarPeriodo() {
		driver.findElement(By.xpath("//a[@title='Alterar Período']")).click();
	}	

	//************************Tela Alterar Periodo***********************
	
	public void selecionarPeriodo(String periodo) {
		Select select = new Select(driver.findElement(By.name("refCompetencia")));
		select.selectByValue(periodo);
	}
	public void clicarBotaoMesReferenciaContinuar() {
		driver.findElement(By.xpath("//input[@value='Continuar']")).click();
	}		
	//************************Tela Funcionario***********************
	
	public void clicarSelecionarFunc() {
		driver.findElement(By.xpath("//a[@class='chosen-single chosen-default']")).click();
	}		
	
	public void preencherFuncionario(String texto) {
		driver.findElement(By.xpath("//INPUT[@type='text']")).sendKeys(texto + Keys.RETURN);
	}
	
	public List<Ponto> capturarRegistroPonto() {
        
        List<WebElement> tr_collection = driver.findElements(By.xpath("//table[@style='border-top: #000 solid 1px;border-left: #000 solid 1px']/tbody/tr"));
        List<Ponto> pontos = new ArrayList<Ponto>();
        int row_num,col_num;
        row_num=1;
        for(WebElement trElement : tr_collection)
        {
            List<WebElement> td_collection=trElement.findElements(By.xpath("td"));
            col_num=0;
            Ponto novoPonto = new Ponto();
            for(WebElement tdElement : td_collection)
            {
           	 
           	 switch (col_num) {
                case 0:
               	 novoPonto.setData(tdElement.getText());
                    break;
                case 1:
               	 String[] marcacoes = tdElement.getText().split(" -");
               	 List<Horario> horarios = new ArrayList<Horario>();
               	 for(String marcacao : marcacoes){
               		 Horario horario = new Horario();
               		 horario.setHorario(marcacao.replace("+", ""));
               		 horarios.add(horario);
               	 }
               	 while(horarios.size()<6) {
               		 Horario horario = new Horario();
               		 horario.setHorario(" ");
               		 horarios.add(horario);
               	 }
               	 novoPonto.setMarcacao(horarios);
                    break;
                case 2:
               	 novoPonto.setCalculos(tdElement.getText());
                    break;
                case 3:
               	 novoPonto.setObservacoes(tdElement.getText());
                    break;
                default:
               	 break;
             }
                col_num++;
            }
            pontos.add(novoPonto);
            row_num++;
        } 
        
       return pontos;

	}

}

