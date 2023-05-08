import PySimpleGUI as sg
import pyodbc
import os
import shutil
import re

# lista de ODBC, sorted alphabetically
odbc_list = sorted(pyodbc.dataSources())

# cria layout da janela principal
layout = [
	[sg.Text('Selecione o ODBC')],
	[sg.Listbox(values=list(odbc_list), size=(30, 6), key='odbc')],
	[sg.Button('Conectar', key='connect'), sg.Text('')],
	[sg.Text('Número da nota fiscal:'), sg.InputText(key='num_nota', disabled=True)],
	[sg.Text('Data da nota fiscal (DDMMAAAA):'), sg.InputText(key='data_nota', disabled=True)],
	[sg.Submit(disabled=True), sg.Cancel()]
]

# cria a janela principal
window = sg.Window('Consulta de nota fiscal', layout)

# loop para eventos da janela principal
while True:
	event, values = window.read()
	if event == sg.WINDOW_CLOSED or event == 'Cancel':
		break
	if event == 'connect':
		# pega o ODBC selecionado
		selected_odbc = values['odbc'][0]
		# tenta se conectar ao banco de dados com o ODBC selecionado
		try:
			conn = pyodbc.connect('DSN=' + selected_odbc)
			sg.popup('Conexão estabelecida com sucesso!')
			window['connect'].update(disabled=True)
			window['odbc'].update(disabled=True)
		except Exception as e:
			sg.popup_error(f'Erro ao conectar ao banco de dados: {str(e)}')
			conn = None
			window['connect'].update(disabled=False)
		if conn is not None:
			window['num_nota'].update(disabled=False)
			window['data_nota'].update(disabled=False)
			window['Submit'].update(disabled=False)
	if event == 'Submit':
		# pega o ODBC selecionado
		selected_odbc = values['odbc'][0]
		# conecta ao banco de dados com o ODBC selecionado
		conn = pyodbc.connect('DSN=' + selected_odbc)
		# tenta fazer a consulta e mostra o resultado em uma nova janela
		try:
			cursor = conn.cursor()
			cursor.execute(
				"select getdatanormal(nsu004) as Data_nota, nsu003 as Série,nsu005 as Número, replace(replace(nsu006,2,'Saída'),1,'Entrada') as Tipo, nsu004, nsu006, empresa, nsu002, SUBSTRING(nsu011, (CHARINDEX(']', nsu011)-44) , 44) as Chave_NFE from ges_359 where nsu011 like '%dupli%' and nsu011 like '%NF%' and nsu004=(select getdatanumero(?)) and nsu005=?",
				values['data_nota'], values['num_nota'])
			data = cursor.fetchone()
			if data is None:
				sg.popup_error('Nota fiscal não encontrada ou não está duplicada!')
				continue
			verifica_chave = re.findall(r'\d{44}', data[8])
			if verifica_chave:
				verifica_chave = verifica_chave[0]
			else:
				sg.popup_error('Não foi possível encontrar a chave da NFe no NSU011, entre em contato com o suporte!')
				continue
			layout_resultado = [
				[sg.Text('Resultado')],
				[sg.Text(f'Data: {data[0]}')],
				[sg.Text(f'Série: {data[1]}')],
				[sg.Text(f'Número: {data[2]}')],
				[sg.Text(f'Tipo: {data[3]}')],
				[sg.Text(f'Empresa: {data[6]}')],
				[sg.Text(f'Local: {data[7]}')],
				[sg.Button('Confirmar', key='confirmar'), sg.Button('Sair', key='sair')]
			]
			window_resultado = sg.Window('Resultado da consulta', layout_resultado)
			while True:
				event, values = window_resultado.read()
				if event == sg.WINDOW_CLOSED or event == 'sair':
					break
				if event == 'confirmar':
					window_resultado.close()
					layout_protocolo = [[sg.Text('Chave NFE:'), sg.InputText(data[8], key='chave_nfe', disabled=True),
										 sg.Button('Copiar', key='copiar_chave')],
										[sg.Button('Baixar XML', key='baixar_xml')],
										[sg.Text('Inserir arquivo XML:'),
										 sg.Input(key='input_file', visible=False, enable_events=True),
										 sg.FileBrowse('Selecionar arquivo', file_types=(("XML Files", "*.xml"),),
													   target='input_file', key='file_browse', disabled=False)],
										[sg.Text('', key='file_name')],
										[sg.Button('Corrigir NF-e', key='file_insert', disabled=False)],
										]
					window_protocolo = sg.Window('Autorização de nota fiscal', layout_protocolo)

					while True:
						event, values = window_protocolo.read()

						if event == sg.WINDOW_CLOSED:
							break

						if event == 'copiar_chave':
							sg.clipboard_set(values['chave_nfe'])
							sg.popup('Chave NFE copiada para a área de transferência!')

						if event == 'baixar_xml':
							import webbrowser

							webbrowser.open('https://www.nfe.fazenda.gov.br/portal/consultaRecaptcha.aspx')

						if event == 'input_file':
							window_protocolo['file_name'].update(
								value='Arquivo selecionado: ' + os.path.basename(values['input_file']))

						if event == 'file_insert':
							# verifica se foi selecionado um arquivo
							if not values['input_file']:
								sg.popup_error('Selecione um arquivo para inserir!')
								continue

							chave_nfe = values['chave_nfe']

							# abre o arquivo selecionado
							with open(values['input_file'], 'r') as file:
								# le o conteudo do arquivo
								file_content = file.read()

								# parseia o xml e obtém o valor das tags nProt e nNF
								nProt = re.findall(r'<nProt>(\d+)</nProt>', file_content)
								nProt = nProt[0] if nProt else None
								nNF = re.findall(r'<nNF>(\d+)</nNF>', file_content)
								nNF = nNF[0] if nNF else None
								chave_xml = re.findall(r'<chNFe>(\d+)</chNFe>', file_content)
								chave_xml = chave_xml[0] if chave_xml else None
								numero_nf = data[2]

								# verifica se a chave_nfe esta no conteudo do arquivo
								if chave_nfe != chave_xml:
									sg.popup_error(
										f'Chave NFE divergente!\n\nChave da NFE:\n{chave_nfe}\n\n'
										f'Chave no XML:\n'
										f'{chave_xml}')
									continue

								# verifica se o número da nota no XML é o mesmo da chave NFE
								if numero_nf != int(nNF):
									sg.popup_error(
										f'Número da nota divergente!\n\nNúmero da NFE:\n{numero_nf}\n\nNúmero no XML:\n{nNF}')
									# sg.popup(type(numero_nf),type(nNF)) verifica os tipos
									continue

								# verifica se a nota está válida
								xMotivo = re.findall(r'<xMotivo>(.*)</xMotivo>', file_content)
								xMotivo = xMotivo[0] if xMotivo else None
								if xMotivo != 'Autorizado o uso da NF-e':
									sg.popup_error(
										f'A nota não está válida! O status atual da mesma é: {xMotivo}, entre em contato com o suporte!')
									continue
								try:
									cursor.execute(
										"update ges_359 set nsu009=?, nsu007=?, nsu008=1, nsu011='Autorizado o uso da NF-e', nsu015=0, nsu016=0, nsu017=0 where nsu004=? and nsu005=? and nsu011 like '%dupli%' and nsu011 like '%NF%'",
										(nProt, data[8], data[4], data[2])
									)
									conn.commit()
									window_protocolo['file_browse'].update(disabled=False)
									window_protocolo['file_insert'].update(disabled=False)
								except Exception as e:
									sg.popup(f'Erro ao inserir protocolo: {str(e)}')
								file.close()
								output_file_name = 'nfe-' + chave_nfe + '.xml'
								# move o arquivo para a pasta de saída
								if data[5] == '2':
									shutil.move(values['input_file'], r'C:\Abase\Gestor\NFE_Saida\\' + output_file_name)
									window_protocolo.close()
									sg.popup('Nota corrigida com sucesso! Gere a Danfe pelo Gestor.')
								elif data[5] == '1':
									shutil.move(values['input_file'],
												r'C:\Abase\Gestor\NFE_Entrada\\' + output_file_name)
									window_protocolo.close()
									sg.popup('Nota corrigida com sucesso! Gere a Danfe pelo Gestor.')
								else:
									window_protocolo.close()
									sg.popup("Valor de nsu006 inválido.")
						if event == sg.WINDOW_CLOSED or event == 'sair':
							break
					window_protocolo.close()
					break
			window_resultado.close()
		except Exception as e:
			sg.popup_error(f'Erro ao consultar banco de dados: {str(e)}')
		finally:
			conn.close()
# fecha a janela principal e termina o programa
window.close()