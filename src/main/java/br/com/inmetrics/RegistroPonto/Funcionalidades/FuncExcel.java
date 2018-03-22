package br.com.inmetrics.RegistroPonto.Funcionalidades;

import java.io.File;
import java.io.IOException;
import java.sql.Driver;
import java.util.List;

import br.com.inmetrics.RegistroPonto.Apoio.ArquivoExcel;
import br.com.inmetrics.RegistroPonto.Apoio.Ponto;

public class FuncExcel {
	ArquivoExcel excel = new ArquivoExcel();

	public void validarExcel(File arquivo) throws IOException {
		excel.gerarArquivoExcel(arquivo);
	}
	
	public void gravarDadosNoExcel(List<Ponto> pontos, String nome) {
		excel.gerarPlanilhaExcel(nome);
		excel.inserirDadosExcel(pontos);
	}
	
	public void salvarFinalizarExcel(File arquivo) {
		excel.salvarArquivoExcel(arquivo);
	}
	
}
