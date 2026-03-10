#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script para extrair dados do EXTRATO.TXT e CREDITO.xlsx
e gerar clienteData em formato JavaScript
"""

import re
import json
from pathlib import Path
import openpyxl

class ParseadorExtrato:
    def __init__(self, arquivo_txt):
        self.arquivo = arquivo_txt
        self.dados = {}
        self.linhas = []

    def ler_arquivo(self):
        """Lê o arquivo TXT"""
        try:
            # Tenta latin-1 primeiro (mais comum para acentos em português)
            with open(self.arquivo, 'r', encoding='latin-1') as f:
                self.linhas = f.readlines()
            return True
        except Exception as e:
            print(f"❌ Erro ao ler {self.arquivo}: {e}")
            return False

    def extrair_cliente(self):
        """Extrai informações do cliente da primeira seção"""
        cliente = {
            "id": "",
            "nome": "",
            "matricula": "",
            "dataNascimento": "",
            "localAcerto": "",
            "areaPropria": 0.0,
            "areaArrendada": 0.0,
            "procurador": "",
            "procuradorId": "",
            "emissao": "",
            "dataAtualizacao": "",
            "estado": "PR"
        }

        for linha in self.linhas[:20]:
            match = re.match(r'^(\d+)\s+(.+?)\s{2,}Matr.*cula:\s+(\d+)', linha)
            if match:
                cliente["id"] = match.group(1)
                cliente["nome"] = match.group(2).strip()
                cliente["matricula"] = match.group(3)
                break

        for linha in self.linhas[:20]:
            if "Local de Acerto:" in linha:
                match_local = re.search(r'Local de Acerto:\s*\d+\s+(.+?)\s{2,}Dt\.Emiss', linha)
                if match_local:
                    cliente["localAcerto"] = match_local.group(1).strip()
                match_data = re.search(r'Dt\.Emiss.*?:\s*(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}:\d{2})', linha)
                if match_data:
                    d = match_data.group(1).split('.')
                    cliente["emissao"] = f"{d[0]}/{d[1]}/{d[2]}"
                    cliente["dataAtualizacao"] = match_data.group(2)
                break

        for linha in self.linhas[:10]:
            if "reas Pr" in linha:
                match_areas = re.search(r'(?:reas|Áreas)\s+Pr.*?:\s*([\d.]+,\d+)\s+(?:.*?)(?:reas|Áreas)\s+Arrendadas:\s*([\d.]+,\d+)', linha)
                if match_areas:
                    cliente["areaPropria"] = float(match_areas.group(1).replace('.', '').replace(',', '.'))
                    cliente["areaArrendada"] = float(match_areas.group(2).replace('.', '').replace(',', '.'))
                break

        for linha in self.linhas:
            if "Procurador(es):" in linha:
                match_proc = re.search(r'Procurador\(es\):\s*(\d+)\s*-\s*(.+?)(?:\n|$)', linha)
                if match_proc:
                    cliente["procuradorId"] = match_proc.group(1)
                    cliente["procurador"] = match_proc.group(2).strip()
                break

        # Adicionar data de nascimento fictícia (será substituída por N8N depois)
        cliente["dataNascimento"] = "15/08/1972"

        return cliente

    def extrair_debitos(self):
        """Extrai débitos de títulos"""
        debitos = []
        inicio = False
        for linha in self.linhas:
            if "BITOS DE T" in linha:
                inicio = True
                continue
            if not inicio:
                continue
            if "Descri" in linha or not linha.strip():
                continue
            if "RESUMO" in linha:
                break
            match = re.match(r'^(.+?)\s+(\d{2}\.\d{2}\.\d{4})\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})', linha)
            if match and "TOTAL" not in match.group(1):
                debitos.append({
                    "descricao": match.group(1).strip(),
                    "vencimento": match.group(2),
                    "valorNominal": float(match.group(3).replace('.', '').replace(',', '.')),
                    "valorAtual": float(match.group(4).replace('.', '').replace(',', '.'))
                })
        return debitos

    def extrair_debitos_por_mes(self):
        """Extrai resumo de débitos por mês"""
        labels, valores = [], []
        inicio = False
        meses = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                 "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        for linha in self.linhas:
            if "RESUMO" in linha and "BITO" in linha:
                inicio = True
                continue
            if not inicio:
                continue
            if "TOTAL" in linha:
                break
            match = re.match(r'^(\d{2}/\d{4})\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})', linha)
            if match:
                m, a = match.group(1).split('/')
                labels.append(f"{meses[int(m)]}/{a}")
                valores.append(float(match.group(2).replace('.', '').replace(',', '.')))
        return {"labels": labels, "valores": valores}

    def extrair_contratos(self):
        """Extrai contratos de insumos com preço definido"""
        contratos = []
        inicio = False
        for linha in self.linhas:
            if "CONTRATOS DE INSUMOS COM PRE" in linha:
                inicio = True
                continue
            if not inicio:
                continue
            if "MERCADORIAS A RETIRAR CONTRATOS PAGOS" in linha:
                break
            if not linha.strip() or "Pedido" in linha or "Nr. Pedido" in linha or linha.startswith("|") or linha.startswith("-") or "Página" in linha:
                continue

            # Validar que começa com 10 dígitos
            if not re.match(r'^\d{10}', linha):
                continue

            # Formato: NUMERO  DESCRICAO  LOCAL  VENCIMENTO  VALOR  QUANTIDADE  UNIDADE
            # Exemplo: 1009083868 ZAPP QI 620 20L                          PALOTINA          01.04.2026            70.741,81             120,000  UN

            # 1. Número (10 dígitos)
            numero = linha[:10].strip()

            # 2. Vencimento (data DD.MM.YYYY) - está antes da quantidade
            vencimento_match = re.search(r'(\d{2}\.\d{2}\.\d{4})', linha)
            if not vencimento_match:
                continue
            vencimento = vencimento_match.group(1)

            # 3. Unidade (última palavra)
            unidade_match = re.search(r'\s([A-Z]{2,})\s*$', linha)
            if not unidade_match:
                continue
            unidade = unidade_match.group(1)

            # 4. Quantidade (número antes da unidade)
            quantidade_match = re.search(r'([\d,]+)\s+[A-Z]{2,}\s*$', linha)
            if not quantidade_match:
                continue
            quantidade = float(quantidade_match.group(1).replace(',', '.'))

            # 5. Valor (número com ponto de milhar antes da quantidade)
            valor_match = re.search(r'([\d\.]+,\d{2})\s+[\d,]+\s+[A-Z]{2,}', linha)
            if not valor_match:
                continue
            valor_str = valor_match.group(1).replace('.', '').replace(',', '.')
            try:
                valor = float(valor_str)
            except:
                continue

            # 6. Resto (descrição + local) - entre número e vencimento
            resto = re.sub(r'^\d{10}\s+', '', linha)  # Remove número
            resto = re.sub(r'\d{2}\.\d{2}\.\d{4}[\s\S]*$', '', resto)  # Remove vencimento em diante

            # Split por múltiplos espaços
            partes = re.split(r'\s{2,}', resto.strip())
            if len(partes) >= 1:
                descricao = partes[0].strip()
                local = partes[-1].strip() if len(partes) > 1 else ""
            else:
                continue

            contratos.append({
                "numero": numero,
                "descricao": descricao,
                "local": local,
                "vencimento": vencimento,
                "valor": valor,
                "quantidade": quantidade,
                "unidade": unidade
            })
        return contratos

    def extrair_mercadorias(self):
        """Extrai mercadorias a retirar - contratos pagos"""
        mercadorias = []
        inicio = False
        for linha in self.linhas:
            if "MERCADORIAS A RETIRAR CONTRATOS PAGOS" in linha:
                inicio = True
                continue
            if not inicio:
                continue
            if "VENDA FUTURA" in linha or "POSI" in linha:
                break
            if not linha.strip() or "Pedido" in linha or linha.startswith("|") or linha.startswith("-"):
                continue
            # Regex corrigido: número de pedido com 10 dígitos
            match = re.match(r'^(\d{10})\s+(.+?)\s{2,}([A-Z][A-Z\s]+?)\s+([\d]+,\d{3})\s+(\w+)', linha)
            if match:
                mercadorias.append({
                    "numero": match.group(1),
                    "descricao": match.group(2).strip(),
                    "local": match.group(3).strip(),
                    "quantidade": float(match.group(4).replace(',', '.')),
                    "unidade": match.group(5).strip(),
                    "valor": 0
                })
        return mercadorias

    def extrair_mercadorias_venda_futura(self):
        """Extrai mercadorias a retirar - venda futura"""
        mercadorias = []
        inicio = False
        for linha in self.linhas:
            if "VENDA FUTURA" in linha:
                inicio = True
                continue
            if not inicio:
                continue
            if "POS" in linha and "ICMS" in linha:
                break
            if not linha.strip() or "Pedido" in linha or "Descri" in linha or linha.startswith("|") or linha.startswith("-"):
                continue

            # Validar que começa com número de pedido
            if not re.match(r'^\d{10}', linha):
                continue

            # 1. Extrair número (10 dígitos)
            numero = linha[:10].strip()

            # 2. Extrair NF (padrão 9dig-3dig)
            nf_match = re.search(r'(\d{9}-\d{3})', linha)
            if not nf_match:
                continue
            nf = nf_match.group(1)

            # 3. Extrair unidade (última palavra)
            unidade_match = re.search(r'\s([A-Z]{2,})\s*$', linha)
            if not unidade_match:
                continue
            unidade = unidade_match.group(1)

            # 4. Extrair quantidade (número com vírgula, antes da unidade)
            quantidade_match = re.search(r'([\d,]+)\s+[A-Z]{2,}\s*$', linha)
            if not quantidade_match:
                continue
            quantidade = float(quantidade_match.group(1).replace(',', '.'))

            # 5. Remover o final da linha (quantidade + unidade) para processar meio
            linha_sem_final = re.sub(r'([\d,]+)\s+[A-Z]{2,}\s*$', '', linha)

            # 6. Remover o início (número + NF) para ficar só com descrição + local
            linha_sem_inicio = re.sub(r'^\d{10}\s+\d{9}-\d{3}\s+', '', linha_sem_final)

            # 7. Split do meio (descrição e local) por múltiplos espaços
            partes = re.split(r'\s{2,}', linha_sem_inicio.strip())

            if len(partes) == 1:
                # Só uma parte = descrição, sem local
                descricao = partes[0].strip()
                local = ""
            elif len(partes) == 2:
                # Duas partes = descrição e local
                descricao = partes[0].strip()
                local = partes[1].strip()
            else:
                # Mais de 2 = primeira é desc, última é local
                descricao = partes[0].strip()
                local = partes[-1].strip()

            mercadorias.append({
                "numero": numero,
                "nf": nf,
                "descricao": descricao,
                "local": local,
                "quantidade": quantidade,
                "unidade": unidade,
                "valor": 0
            })
        return mercadorias

    def extrair_debitos_em_produtos(self):
        """Extrai débitos em produtos da seção 'DÉBITOS EM PRODUTO'"""
        debitos = []
        inicio = False
        for linha in self.linhas:
            # Procurar por "EM PRODUTO" (não "DÉBITOS" pois pode ter encoding diferente)
            if "EM PRODUTO" in linha and "BITOS" in linha:
                inicio = True
                continue
            if not inicio:
                continue
            if "EM T" in linha and "TULO" in linha:  # DÉBITOS DE TÍTULOS
                break
            if not linha.strip() or "Descri" in linha or "Contrato" in linha or linha.startswith("|") or linha.startswith("-"):
                continue

            # Padrão: DESCRICAO  CONTRATO(7dig)  PRODUTO  QUANTIDADE  VENCIMENTO
            # Exemplo: CÉDULA PROD. RURAL  3470364            SOJA COMERCIAL  27.649,000    01.04.2026

            # Procurar: número de contrato (7 dígitos) seguido de espaços
            contrato_match = re.search(r'(\d{7})\s+', linha)
            if not contrato_match:
                continue

            contrato = contrato_match.group(1)
            contrato_pos = contrato_match.start()

            # Extrair descrição (tudo antes do contrato)
            descricao = linha[:contrato_pos].strip()

            # Extrair resto (após contrato)
            resto = linha[contrato_match.end():]

            # Do resto, extrair: PRODUTO ... QUANTIDADE ... VENCIMENTO
            # Vencimento é data no final (DD.MM.YYYY)
            vencimento_match = re.search(r'(\d{2}\.\d{2}\.\d{4})\s*$', resto)
            if not vencimento_match:
                continue
            vencimento = vencimento_match.group(1)

            # Quantidade é o número antes da vencimento
            quantidade_match = re.search(r'([\d\.]+,\d+)\s+\d{2}\.\d{2}\.\d{4}', resto)
            if not quantidade_match:
                continue
            quantidade_str = quantidade_match.group(1).replace('.', '').replace(',', '.')
            try:
                quantidade = float(quantidade_str)
            except:
                continue

            # Produto é o que fica entre contrato e quantidade
            produto_parte = re.sub(r'([\d\.]+,\d+)\s+\d{2}\.\d{2}\.\d{4}\s*$', '', resto)
            produto = produto_parte.strip()

            debitos.append({
                "descricao": descricao,
                "contrato": contrato,
                "produto": produto,
                "quantidade": quantidade,
                "vencimento": vencimento
            })
        return debitos

    def extrair_icms(self):
        """Extrai posição de ICMS"""
        icms = {
            "apresentadas": 0,
            "reconhecido": 0,
            "homologacao": 0,
            "aguardandoLiberacao": 0,
            "homologados": 0,
            "totalCreditosPresentados": 0
        }
        inicio = False
        for linha in self.linhas:
            if "POS" in linha and "ICMS" in linha:
                inicio = True
                continue
            if not inicio:
                continue
            if "TOTAL DOS" in linha:
                match = re.search(r'R\$\s+([\d.]+,\d{2})', linha)
                if match:
                    icms["totalCreditosPresentados"] = float(match.group(1).replace('.', '').replace(',', '.'))
                break
            for chave, padrao in [
                ("apresentadas", "Apresentadas"),
                ("reconhecido", "Reconhecido"),
                ("homologacao", "Em Homologa"),
                ("aguardandoLiberacao", "aguardando libera"),
                ("homologados", "Homologados")
            ]:
                if padrao in linha:
                    match = re.search(r'R\$\s+([\d.]+,\d{2})', linha)
                    if match:
                        icms[chave] = float(match.group(1).replace('.', '').replace(',', '.'))
        return icms

    def extrair_todos(self):
        """Extrai todos os dados do TXT"""
        if not self.ler_arquivo():
            return None
        return {
            "cliente": self.extrair_cliente(),
            "debitosTitulos": self.extrair_debitos(),
            "debitosPorMes": self.extrair_debitos_por_mes(),
            "contratos": self.extrair_contratos(),
            "mercadorias": self.extrair_mercadorias(),
            "mercadoriasVendaFutura": self.extrair_mercadorias_venda_futura(),
            "debitosEmProdutos": self.extrair_debitos_em_produtos(),
            "creditoICMS": self.extrair_icms()
        }


class ParseadorCredito:
    def __init__(self, arquivo_xlsx):
        self.arquivo = arquivo_xlsx

    def extrair_segmentos(self):
        """Extrai segmentos de crédito do XLSX"""
        segmentos = []
        try:
            wb = openpyxl.load_workbook(self.arquivo, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=5, values_only=True):
                # Ignora linhas sem parceiro ou sem segmento (linha de totais)
                if not row[0] or not row[2] or row[2] == '':
                    continue
                segmentos.append({
                    "segmento": str(row[2]),
                    "descricao": str(row[7]) if row[7] else "",
                    "limite": float(row[3]) if row[3] else 0,
                    "compromisso": float(row[4]) if row[4] else 0,
                    "excesso": float(row[5]) if row[5] else 0,
                    "esgotamento": float(row[6]) if row[6] else 0
                })
            wb.close()
            return segmentos
        except Exception as e:
            print(f"❌ Erro ao ler CREDITO.xlsx: {e}")
            return []


def gerar_javascript(dados_txt, segmentos_credito):
    """Monta o objeto clienteData em JavaScript"""
    from datetime import datetime

    cliente = dados_txt.get("cliente", {})
    debitos = dados_txt.get("debitosTitulos", [])
    debitos_mes = dados_txt.get("debitosPorMes", {})
    contratos = dados_txt.get("contratos", [])
    mercadorias = dados_txt.get("mercadorias", [])
    mercadorias_venda_futura = dados_txt.get("mercadoriasVendaFutura", [])
    debitos_em_produtos = dados_txt.get("debitosEmProdutos", [])
    icms = dados_txt.get("creditoICMS", {})

    total_debitos = sum(d.get("valorAtual", 0) for d in debitos)

    # Calcular débitos vencidos (vencimento < hoje)
    hoje = datetime.now().date()
    debitos_vencidos_valor = 0
    for d in debitos:
        try:
            # Converter DD.MM.YYYY para data
            partes = d.get("vencimento", "").split(".")
            if len(partes) == 3:
                vencto = datetime(int(partes[2]), int(partes[1]), int(partes[0])).date()
                if vencto < hoje:
                    debitos_vencidos_valor += d.get("valorAtual", 0)
        except:
            pass

    area_total = cliente.get("areaPropria", 0) + cliente.get("areaArrendada", 0)

    js_content = f'''// ==DATA:START==
const clienteData = {{
    cliente: {{
        id: "{cliente.get('id', '')}",
        nome: "{cliente.get('nome', '')}",
        matricula: "{cliente.get('matricula', '')}",
        dataNascimento: "{cliente.get('dataNascimento', '')}",
        localAcerto: "{cliente.get('localAcerto', '')}",
        areaPropria: {cliente.get('areaPropria', 0)},
        areaArrendada: {cliente.get('areaArrendada', 0)},
        procurador: "{cliente.get('procurador', '')}",
        procuradorId: "{cliente.get('procuradorId', '')}",
        emissao: "{cliente.get('emissao', '')}",
        dataAtualizacao: "{cliente.get('dataAtualizacao', '')}",
        estado: "{cliente.get('estado', 'PR')}"
    }},
    classificacoes: [
        /* FONTE: C4C Cloud OData - CorporateAccountCollection */
        /* Dados provisórios - Serão atualizados automaticamente pelo N8N */
        {{ nome: "Perfil Comportamental", valor: "Criativo" }},
        {{ nome: "Negociação e Atendimento", valor: "-" }},
        {{ nome: "Perfil de Compra", valor: "-" }},
        {{ nome: "Perfil Negócio", valor: "-" }},
        {{ nome: "Perfil Preço (Compras)", valor: "30" }},
        {{ nome: "Perfil Negócio (Compras)", valor: "50" }},
        {{ nome: "Perfil Relacionamento (Compras)", valor: "20" }}
    ],
    historico: [
        /* FONTE: C4C Cloud OData - ActivityCollection */
        /* Dados de Exemplo - Serão atualizados automaticamente pelo N8N */
        {{ data: "05/03/2026", hora: "14:30", consultor: "Carlos Mendes", tipo: "Visita Técnica", assunto: "Análise de aplicação ZAPP QI 620 - Controle de plantas daninhas", observacoes: "Cliente interessado em aumentar volume de compra. Recomendado treinar aplicadores." }},
        {{ data: "02/03/2026", hora: "10:15", consultor: "Marina Silva", tipo: "Telefonema", assunto: "Seguimento de pedido de fertilizante FERT 02 20 18", observacoes: "Confirmado entrega para próxima semana. Cliente satisfeito." }},
        {{ data: "26/02/2026", hora: "09:00", consultor: "Roberto Santos", tipo: "Visita Comercial", assunto: "Apresentação de novas sementes SEM SOJA M7739 IPRO", observacoes: "Cliente demonstrou interesse. Solicitou amostra para teste em 2 hectares." }},
        {{ data: "20/02/2026", hora: "15:45", consultor: "Carlos Mendes", tipo: "Telefonema", assunto: "Resolução de problema com vencimento de títulos", observacoes: "Documentação corrigida. Novo vencimento: 16/01/2026. Débito quitado." }},
        {{ data: "14/02/2026", hora: "11:20", consultor: "Juliana Costa", tipo: "Visita Técnica", assunto: "Diagnóstico de deficiência nutricional em soja", observacoes: "Recomendado aplicação de fertilizante foliar. Planejado novo atendimento em 15 dias." }}
    ],
    oportunidades: [
        /* FONTE: C4C Cloud OData - OpportunityCollection */
        /* Dados de Exemplo - Serão atualizados automaticamente pelo N8N */
        {{ status: "aberta", titulo: "Aquisição de ZAPP QI 620 - Grande Volume", descricao: "Cliente interessado em aumentar volume anual de 513 para 800 unidades de ZAPP QI 620. Demanda para próxima safra 2026/2027.", valor: 145000.00, dataAbertura: "05/03/2026", responsavel: "Carlos Mendes", icone: "package" }},
        {{ status: "em-andamento", titulo: "Teste de Sementes SEM SOJA M7739 IPRO", descricao: "Amostra de 2 hectares em análise. Cliente aguardando resultados de produtividade. Próxima visita 30/03/2026.", valor: 28000.00, dataAbertura: "26/02/2026", responsavel: "Roberto Santos", icone: "leaf" }},
        {{ status: "aberta", titulo: "Contrato Anual de Fertilizantes 2026", descricao: "Proposta de contrato anual para FERT 02 20 18 TMIX EVOLUTION. Visando estabilizar preço e volume. Negociação em fase inicial.", valor: 85000.00, dataAbertura: "01/03/2026", responsavel: "Marina Silva", icone: "trending-up" }}
    ],
    kpis: {{
        totalDebitos: {round(total_debitos, 2)},
        debitosVencidos: {round(debitos_vencidos_valor, 2)},
        areaTotal: {area_total},
        dataAtualizacao: "{cliente.get('emissao', '')}"
    }},
    debitosTitulos: [
'''
    for i, d in enumerate(debitos):
        virgula = ',' if i < len(debitos) - 1 else ''
        js_content += f'        {{ descricao: "{d.get("descricao","")}", vencimento: "{d.get("vencimento","")}", valorNominal: {d.get("valorNominal",0)}, valorAtual: {d.get("valorAtual",0)} }}{virgula}\n'

    js_content += '    ],\n    contratos: [\n'

    for i, c in enumerate(contratos):
        virgula = ',' if i < len(contratos) - 1 else ''
        js_content += f'        {{ numero: "{c.get("numero","")}", descricao: "{c.get("descricao","")}", local: "{c.get("local","")}", vencimento: "{c.get("vencimento","")}", valor: {c.get("valor",0)}, quantidade: {c.get("quantidade",0)}, unidade: "{c.get("unidade","")}" }}{virgula}\n'

    js_content += '    ],\n    mercadorias: [\n'

    for i, m in enumerate(mercadorias):
        virgula = ',' if i < len(mercadorias) - 1 else ''
        js_content += f'        {{ numero: "{m.get("numero","")}", descricao: "{m.get("descricao","")}", local: "{m.get("local","")}", quantidade: {m.get("quantidade",0)}, unidade: "{m.get("unidade","")}", valor: {m.get("valor",0)} }}{virgula}\n'

    js_content += f'''    ],
    creditoICMS: {{
        apresentadas: {icms.get('apresentadas', 0)},
        reconhecido: {icms.get('reconhecido', 0)},
        homologacao: {icms.get('homologacao', 0)},
        aguardandoLiberacao: {icms.get('aguardandoLiberacao', 0)},
        homologados: {icms.get('homologados', 0)},
        totalCreditosPresentados: {icms.get('totalCreditosPresentados', 0)}
    }},
    creditoSegmentos: [
'''
    for i, s in enumerate(segmentos_credito):
        virgula = ',' if i < len(segmentos_credito) - 1 else ''
        js_content += f'        {{ segmento: "{s.get("segmento","")}", descricao: "{s.get("descricao","")}", limite: {s.get("limite",0)}, compromisso: {s.get("compromisso",0)}, excesso: {s.get("excesso",0)}, esgotamento: {s.get("esgotamento",0)} }}{virgula}\n'

    js_content += f'''    ],
    debitosPorMes: {{
        labels: {json.dumps(debitos_mes.get('labels', []), ensure_ascii=False)},
        valores: {json.dumps(debitos_mes.get('valores', []))}
    }},
    mercadoriasVendaFutura: [
'''

    # Adiciona mercadorias de venda futura se existirem
    for i, m in enumerate(mercadorias_venda_futura):
        virgula = ',' if i < len(mercadorias_venda_futura) - 1 else ''
        js_content += f'        {{ numero: "{m.get("numero","")}", descricao: "{m.get("descricao","")}", local: "{m.get("local","")}", quantidade: {m.get("quantidade",0)}, unidade: "{m.get("unidade","")}", valor: {m.get("valor",0)} }}{virgula}\n'

    js_content += '''    ],
    debitosEmProdutos: [
'''

    # Adiciona débitos em produtos se existirem
    for i, d in enumerate(debitos_em_produtos):
        virgula = ',' if i < len(debitos_em_produtos) - 1 else ''
        js_content += f'        {{ descricao: "{d.get("descricao","")}", contrato: "{d.get("contrato","")}", produto: "{d.get("produto","")}", quantidade: {d.get("quantidade",0)}, vencimento: "{d.get("vencimento","")}" }}{virgula}\n'

    js_content += '''    ]
};
// ==DATA:END==
'''
    return js_content


def main():
    pasta = Path(__file__).parent  # Usa sempre a pasta onde o script está

    print("[PARSER] Parser de Dados - EXTRATO.TXT + CREDITO.xlsx")
    print("=" * 60)

    # 1. Parsear EXTRATO.TXT
    print("\n[1] Extraindo dados do EXTRATO.TXT...")
    parseador_txt = ParseadorExtrato(str(pasta / "EXTRATO.TXT"))
    dados_txt = parseador_txt.extrair_todos()

    if not dados_txt:
        print("[ERRO] Falha ao extrair dados do EXTRATO.TXT")
        return

    print("[OK] Cliente: " + dados_txt['cliente'].get('nome', 'N/A'))
    print("[OK] Débitos: " + str(len(dados_txt['debitosTitulos'])) + " títulos")
    print("[OK] Mercadorias: " + str(len(dados_txt['mercadorias'])) + " itens")

    # 2. Parsear CREDITO.xlsx
    print("\n[2] Extraindo dados do CREDITO.xlsx...")
    parseador_xlsx = ParseadorCredito(str(pasta / "CREDITO.xlsx"))
    segmentos_credito = parseador_xlsx.extrair_segmentos()
    print("[OK] Segmentos: " + str(len(segmentos_credito)) + " segmentos extraídos")

    # 3. Gerar JavaScript
    print("\n[3] Gerando clienteData em JavaScript...")
    js_content = gerar_javascript(dados_txt, segmentos_credito)

    # 4. Salvar clienteData_gerado.js
    arquivo_saida = pasta / "clienteData_gerado.js"
    with open(arquivo_saida, 'w', encoding='utf-8') as f:
        f.write(js_content)
    print("[OK] Arquivo salvo: " + str(arquivo_saida))

    # 5. Injetar no dashboard-cliente-v2.html
    print("\n[4] Injetando dados no dashboard-cliente-v2.html...")
    arquivo_html = pasta / "dashboard-cliente-v2.html"
    try:
        with open(arquivo_html, 'r', encoding='utf-8') as f:
            html_content = f.read()

        # Substitui o bloco entre as tags (simples e robusto)
        pattern = r'// ==DATA:START==.*?// ==DATA:END=='
        novo_bloco = js_content.strip()
        html_novo = re.sub(pattern, novo_bloco, html_content, flags=re.DOTALL)

        if html_novo == html_content:
            print("[AVISO] Dados já estão atualizados no HTML - nenhuma mudança necessária.")
        else:
            with open(arquivo_html, 'w', encoding='utf-8') as f:
                f.write(html_novo)
            print("[OK] HTML atualizado com sucesso!")

    except FileNotFoundError:
        print(f"[AVISO] {arquivo_html} não encontrado. Apenas o .js foi gerado.")

    # 6. Resumo
    print("\n" + "=" * 60)
    print("[SUCESSO] Processo concluído!")
    print(f"  Cliente       : {dados_txt['cliente'].get('nome', 'N/A')}")
    print(f"  Débitos       : {len(dados_txt['debitosTitulos'])} títulos")
    print(f"  Contratos     : {len(dados_txt['contratos'])} itens")
    print(f"  Mercad.Pagas  : {len(dados_txt['mercadorias'])} itens")
    print(f"  Mercad.Futura : {len(dados_txt['mercadoriasVendaFutura'])} itens")
    print(f"  Débitos Prod. : {len(dados_txt['debitosEmProdutos'])} itens")
    print(f"  Segmentos     : {len(segmentos_credito)}")
    print("=" * 60)


if __name__ == "__main__":
    main()
