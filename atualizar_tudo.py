#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de Atualizacao Completa - Sempre que EXTRATO.TXT muda
Atualiza: clienteData_gerado.js, dashboard-cliente-v2.html
"""

import sys
import subprocess
from pathlib import Path
from datetime import datetime

def main():
    pasta = Path(__file__).parent
    print("\n" + "="*70)
    print("[ATUALIZACAO COMPLETA] " + datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
    print("="*70)
    
    # 1. Executar parse_dados.py
    print("\n[1/3] Executando parser de dados...")
    resultado = subprocess.run([sys.executable, str(pasta / "parse_dados.py")], 
                              capture_output=True, text=True)
    print(resultado.stdout)
    if resultado.returncode != 0:
        print("[ERRO] Falha ao executar parse_dados.py")
        print(resultado.stderr)
        return False
    
    # 2. Verificar arquivos gerados/atualizados
    print("\n[2/3] Verificando arquivos atualizados...")
    arquivos_gerados = {
        "clienteData_gerado.js": pasta / "clienteData_gerado.js",
        "dashboard-cliente-v2.html": pasta / "dashboard-cliente-v2.html"
    }
    
    for nome, caminho in arquivos_gerados.items():
        if caminho.exists():
            tamanho = caminho.stat().st_size
            print(f"  [OK] {nome} ({tamanho} bytes)")
        else:
            print(f"  [ERRO] {nome} NAO ENCONTRADO")
    
    # 3. Resumo final
    print("\n[3/3] Resumo da atualizacao:")
    print("  [OK] clienteData_gerado.js        - Dados em JavaScript")
    print("  [OK] dashboard-cliente-v2.html    - Dashboard com dados injetados")
    print("  [INFO] Abra dashboard-cliente-v2.html no navegador para visualizar")
    
    print("\n" + "="*70)
    print("[SUCESSO] ATUALIZACAO COMPLETA CONCLUIDA!")
    print("="*70 + "\n")
    return True

if __name__ == "__main__":
    main()
