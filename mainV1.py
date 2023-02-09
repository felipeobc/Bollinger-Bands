from datetime import date
from openpyxl.chart import Reference
from openpyxl.styles import PatternFill, Font, Alignment
from classes import LeitorAcoes, GerenciadorPlanilha, PropriedadesSerieGradicos

try:
    # acao = input("QUal o codigo da Ação que você quer processar: ").upper()

    acao = "BIDI4"

    leitor_acoes = LeitorAcoes(caminho_arquivo='./dados/')
    leitor_acoes.processa_arquivo(acao)

    gerenciador = GerenciadorPlanilha()

    planilha_dados = gerenciador.adiciona_planilha(titulo_planilha="Dados")

    gerenciador.adiciona_linas(["DATA", "COTAÇÂO", "BANDA INFERIOR", "BANDA SUPERIOR"])

    indice = 2

    for linha in leitor_acoes.dados:
        # DATA
        # 2018-05-10 21:00:00;0.9969
        ano_mes_dia = linha[0].split(" ")[0]
        data = date(
            year=int(ano_mes_dia.split("-")[0]),
            month=int(ano_mes_dia.split("-")[1]),
            day=int(ano_mes_dia.split("-")[2])
        )
        # COTAÇÃO
        cotacao = float(linha[1])

        # Atualiza celulas
        formula_bb_inferior = f'AVERANGE(B{indice}:B{indice + 19}) - 2*STDEV(B{indice}:B{indice + 19})'
        formula_bb_superior = f'AVERANGE(B{indice}:B{indice + 19}) + 2*STDEV(B{indice}:B{indice + 19})'
        gerenciador.atualiza_celula(celula=f'A{indice}', dado=data)
        gerenciador.atualiza_celula(celula=f'B{indice}', dado=cotacao)
        gerenciador.atualiza_celula(celula=f'C{indice}', dado=formula_bb_inferior)
        gerenciador.atualiza_celula(celula=f'D{indice}', dado=formula_bb_superior)

        indice += 1

    gerenciador.adiciona_planilha(titulo_planilha='Grafico')
    gerenciador.mescla_celula(celula_inicio='A1', celula_fim='T2')

    gerenciador.aplica_estilos(celula='A1',
                               estilos=[
                                   ('font', Font(b=True, sz=18, color="FFFFFF")),
                                   ('fill', PatternFill("solid", fgColor="07838F")),
                                   ('alignment', Alignment(vertical="center", horizontal="center"))
                               ])

    gerenciador.atualiza_celula('A1', 'Historico de Cotação')

    referencia_cotacao = Reference(planilha_dados, min_col=2, min_row=2, max_col=4, max_row=indice)
    referencia_datas = Reference(planilha_dados, min_col=1, min_row=2, max_col=1, max_row=indice)

    gerenciador.adiciona_grafico_linha(
        celula='A3',
        comprimento=33.87,
        altura=14.82,
        titulo=f'Cotações - {acao}',
        titulo_eixo_x="Data da Cotação",
        titulo_eixo_y="Valor da Cotação",
        referencia_eixo_x=referencia_cotacao,
        referencia_eixo_y=referencia_datas,
        propriedade_grafico=[
            PropriedadesSerieGradicos(grossura=0, cor_preenchimento='0a55ab'),
            PropriedadesSerieGradicos(grossura=0, cor_preenchimento='a61588'),
            PropriedadesSerieGradicos(grossura=0, cor_preenchimento='12a154')
        ]

    )

    gerenciador.mescla_celula(celula_inicio='I32', celula_fim='L35')
    gerenciador.adiciona_imagem(celula='I32', caminho_imagem='./recursos/b3.png')

    gerenciador.salva_arquivo("./saida/Planilha.xlsx")

except FileNotFoundError:
    print('Cota nao encontrada')
except ValueError:
    print('Formado de dados incorretos')
except AttributeError:
    print('Atributo inexistente')
except Exception as exececao:
    print(f'Ocorreu um erro na execução do programa {str(exececao)}')