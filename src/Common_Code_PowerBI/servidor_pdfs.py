# -*- coding: utf-8 -*-
r"""Serve data/pdfs em http://localhost:8000 para o botão "Abrir licença".

PORQUE É PRECISO
----------------
O Power BI **não abre ficheiros locais** a partir de um botão. Um endereço
`file:///...` é simplesmente ignorado — é uma limitação da própria ferramenta,
não uma configuração que falte. Só aceita `http://` e `https://`.

Depois da mudança para o SharePoint isto deixa de ser preciso: os PDFs passam
a ter endereços https:// reais. Até lá, este servidor dá-lhes um endereço
http:// nesta máquina, e o botão funciona.

COMO USAR
---------
1. Duplo-clique em  Automation files/SERVIDOR_PDFS.bat   (deixar a janela aberta)
2. Abrir o dashboard e usar o botão normalmente.

Fechar a janela desliga o servidor e o botão deixa de abrir os PDFs.

SEGURANÇA
---------
Escuta APENAS em 127.0.0.1 (esta máquina). Ninguém na rede consegue ligar-se
nem ver o arquivo de licenças. Não usar 0.0.0.0 aqui: isso publicaria todas as
licenças na rede da empresa, sem autenticação nenhuma.
"""
from __future__ import annotations

import functools
import http.server
import os
import socket
import socketserver
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
PASTA = os.getenv('EPAL_PDF_DIR', os.path.join(ROOT, 'data', 'pdfs'))
PORTA = int(os.getenv('EPAL_PDF_PORT', '8000'))
ENDERECO = '127.0.0.1'          # NUNCA 0.0.0.0 - ver nota de segurança acima


class Silencioso(http.server.SimpleHTTPRequestHandler):
    """Só regista o que interessa: um pedido por licença aberta."""

    def log_message(self, fmt, *args):        # noqa: A003
        if args and str(args[0]).startswith('GET'):
            alvo = str(args[0]).split(' ')[1] if ' ' in str(args[0]) else str(args[0])
            print(f'  {self.log_date_time_string()}  {alvo[-70:]}')

    def end_headers(self):
        # o Power BI abre o PDF no browser; sem isto alguns browsers descarregam
        self.send_header('Cache-Control', 'no-store')
        super().end_headers()


def porta_livre(porta: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex((ENDERECO, porta)) != 0


def main() -> int:
    if not os.path.isdir(PASTA):
        print(f'ERRO: nao encontrei a pasta dos PDFs:\n  {PASTA}')
        return 1
    porta = PORTA
    if not porta_livre(porta):
        print(f'A porta {porta} ja esta ocupada.')
        for tentativa in range(porta + 1, porta + 10):
            if porta_livre(tentativa):
                porta = tentativa
                break
        else:
            print('Nao encontrei nenhuma porta livre. Feche o outro servidor.')
            return 1
        print(f'A usar a porta {porta} em vez disso.')
        print(f'ATENCAO: o feed aponta para a porta {PORTA}. Para usar a {porta},')
        print(f'         defina EPAL_PDF_PORT={porta} e volte a correr make_powerbi_all.py')

    handler = functools.partial(Silencioso, directory=PASTA)
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer((ENDERECO, porta), handler) as httpd:
        n = sum(len(fs) for _, _, fs in os.walk(PASTA))
        print('=' * 66)
        print(' EPAL - servidor local dos PDFs das licencas')
        print('=' * 66)
        print(f'  a servir : {PASTA}')
        print(f'  em       : http://{ENDERECO}:{porta}/   ({n} ficheiros)')
        print(f'  so nesta maquina - ninguem na rede consegue aceder.')
        print()
        print('  DEIXE ESTA JANELA ABERTA enquanto usa o dashboard.')
        print('  Ctrl+C (ou fechar a janela) desliga o servidor.')
        print('=' * 66)
        print()
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print('\nServidor desligado.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
