package br.com.inmetrics.RegistroPonto.Apoio;

import java.util.List;

public class Ponto {

	private String data;
	private List<Horario> marcacao;
	private String calculos;
	private String observacoes;
	public String getData() {
		return data;
	}
	public void setData(String data) {
		this.data = data;
	}
	public List<Horario> getMarcacao() {
		return marcacao;
	}
	public void setMarcacao(List<Horario> marcacao) {
		this.marcacao = marcacao;
	}
	public String getCalculos() {
		return calculos;
	}
	public void setCalculos(String calculos) {
		this.calculos = calculos;
	}
	public String getObservacoes() {
		return observacoes;
	}
	public void setObservacoes(String observacoes) {
		this.observacoes = observacoes;
	}
}
