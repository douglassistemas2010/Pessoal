#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script para injetar clienteData_gerado.js no dashboard-cliente-v2.html
Substitui o conteúdo entre // ==DATA:START== e // ==DATA:END==
"""

import re
from pathlib import Path

def injetar_dados_no_html():
    """Injeta dados gerados no HTML"""

    pasta = Path("c:/CVale_Desenv/Relatorio_atedimento_previsita")
    arquivo_html = pasta / "dashboard-cliente-v2.html"
    arquivo_js = pasta / "clienteData_gerado.js"

    print("[INJETAR] Sistema de Injeção de Dados")
    print("=" * 60)

    # 1. Verificar se arquivos existem
    print("\n[1] Verificando arquivos...")
    if not arquivo_html.exists():
        print("[ERRO] Arquivo HTML não encontrado: " + str(arquivo_html))
        return False
    print("[OK] HTML encontrado")

    if not arquivo_js.exists():
        print("[ERRO] Arquivo JS não encontrado: " + str(arquivo_js))
        return False
    print("[OK] JS gerado encontrado")

    # 2. Ler arquivo JavaScript
    print("\n[2] Lendo clienteData_gerado.js...")
    try:
        with open(arquivo_js, 'r', encoding='utf-8') as f:
            conteudo_js = f.read()
        print("[OK] Arquivo JS lido: " + str(len(conteudo_js)) + " caracteres")
    except Exception as e:
        print("[ERRO] Falha ao ler JS: " + str(e))
        return False

    # 3. Ler arquivo HTML
    print("\n[3] Lendo dashboard-cliente-v2.html...")
    try:
        with open(arquivo_html, 'r', encoding='utf-8') as f:
            conteudo_html = f.read()
        print("[OK] Arquivo HTML lido: " + str(len(conteudo_html)) + " caracteres")
    except Exception as e:
        print("[ERRO] Falha ao ler HTML: " + str(e))
        return False

    # 4. Procurar pelas tags de delimitação
    print("\n[4] Procurando tags de delimitação...")

    # Padrão: // ==DATA:START== ... // ==DATA:END==
    # Usar DOTALL para capturar quebras de linha
    padrao = r'//\s*==DATA:START==.*?//\s*==DATA:END=='

    if not re.search(padrao, conteudo_html, re.DOTALL):
        print("[ERRO] Tags de delimitação nao encontradas no HTML!")
        print("       Procure por: // ==DATA:START== e // ==DATA:END==")
        return False

    print("[OK] Tags encontradas")

    # 5. Fazer a substituição
    print("\n[5] Substituindo conteúdo...")
    try:
        conteudo_html_novo = re.sub(
            padrao,
            conteudo_js.strip(),
            conteudo_html,
            flags=re.DOTALL
        )
        print("[OK] Substituição realizada")
    except Exception as e:
        print("[ERRO] Falha na substituição: " + str(e))
        return False

    # 6. Verificar se mudou
    if conteudo_html == conteudo_html_novo:
        print("[AVISO] Conteudo HTML nao foi alterado!")
        return False

    # 7. Salvar HTML atualizado
    print("\n[6] Salvando HTML atualizado...")
    try:
        with open(arquivo_html, 'w', encoding='utf-8') as f:
            f.write(conteudo_html_novo)
        print("[OK] HTML salvo com sucesso!")
    except Exception as e:
        print("[ERRO] Falha ao salvar HTML: " + str(e))
        return False

    # 8. Resumo
    print("\n" + "=" * 60)
    print("[SUCESSO] Dados injetados no HTML com sucesso!")
    print("[ARQUIVO] " + str(arquivo_html))
    print("[TAMANHO] HTML agora tem " + str(len(conteudo_html_novo)) + " caracteres")
    print("=" * 60)
    print("\n[PROX-PASSO] Abra o HTML no navegador para verificar!")

    return True


if __name__ == "__main__":
    sucesso = injetar_dados_no_html()
    exit(0 if sucesso else 1)
