package br.com.inmetrics.RegistroPonto.Funcionalidades;

import java.util.ArrayList;
import java.util.List;

import org.openqa.selenium.WebDriver;

import br.com.inmetrics.RegistroPonto.Apoio.ArquivoExcel;
import br.com.inmetrics.RegistroPonto.Apoio.Navegador;
import br.com.inmetrics.RegistroPonto.Apoio.Ponto;
import br.com.inmetrics.RegistroPonto.Telas.Zeus;

public class FuncZeus {
	private Navegador navegador = new Navegador();
	private Zeus zeus;
	private WebDriver driver;
	
	public FuncZeus(WebDriver driver) {
		this.driver = driver;
		zeus = new Zeus(driver);
	}
	
	public void realizarLogin(String usuario, String senha) {
		zeus.preencherUsuario(usuario);
		zeus.preencherSenha(senha);
		zeus.clicarBotaoContinuar();
	}
	
	public void selecionarPeriodo(String periodo) {
		zeus.clicarBotaoAlterarPeriodo();
		zeus.selecionarPeriodo(periodo);
		zeus.clicarBotaoMesReferenciaContinuar();
		zeus.clicarBotaoFrequencia();
	}
	
	public void selecionarFuncionario(String funcionario) {
		zeus.clicarSelecionarFunc();
		zeus.preencherFuncionario(funcionario);
	}
	
	public List<Ponto> capturarRegistrosPonto() {
		return zeus.capturarRegistroPonto();
	}
}
