# -*- coding: utf-8 -*-
r"""Os mapas do dashboard conseguem chegar aos servidores?

PORQUE E QUE ISTO EXISTE
------------------------
Um mapa TOTALMENTE BRANCO no Power BI quase nunca e um problema de dados: o
visual abriu, mas as imagens do mapa ("tiles") nao chegaram. Isso e rede.

O dashboard usa o visual **Azure Maps**, que vai buscar os tiles a
atlas.microsoft.com. O visual de mapa NORMAL do Power BI usa o Bing. Testar os
dois separa tres causas que se parecem todas com "o mapa nao aparece":

  * os dois falham          -> nao ha internet, ou o proxy exige autenticacao;
  * so o Azure Maps falha   -> o proxy deixa passar o Bing mas nao o Azure;
  * os dois respondem       -> a rede esta bem; o problema e a definicao do
                               visual Azure Maps (administrador do Power BI).

Este teste NAO precisa de credenciais nenhumas.
"""
from __future__ import annotations

import ssl
import sys
import urllib.error
import urllib.request

# Alguns proxys de empresa apresentam um certificado cujo campo
# "Basic Constraints" nao esta marcado como critico. O Python 3.13 rejeita-o
# por omissao; os browsers aceitam-no. Relaxamos SO essa regra - a cadeia e o
# nome do servidor continuam a ser verificados.
SITIOS = [
    ('Azure Maps (o que este dashboard usa)', 'https://atlas.microsoft.com'),
    ('Bing Maps (mapa normal do Power BI)', 'https://www.bing.com'),
    ('Power BI (servico)', 'https://app.powerbi.com'),
]


def _contexto() -> ssl.SSLContext:
    ctx = ssl.create_default_context()
    ctx.verify_flags &= ~ssl.VERIFY_X509_STRICT
    return ctx


def testar(url: str, ctx) -> tuple[bool, str]:
    try:
        r = urllib.request.urlopen(url, timeout=10, context=ctx)
        return True, 'HTTP %s' % r.status
    except urllib.error.HTTPError as exc:
        # 401/403 continua a provar que a ligacao chegou la
        return True, 'HTTP %s (respondeu)' % exc.code
    except Exception as exc:
        return False, '%s: %s' % (type(exc).__name__, str(exc)[:60])


def main() -> int:
    ctx = _contexto()
    resultados = []
    for nome, url in SITIOS:
        ok, detalhe = testar(url, ctx)
        resultados.append((nome, ok))
        print('      %-40s %-4s %s' % (nome, 'OK' if ok else 'FALHA', detalhe))

    azure = resultados[0][1]
    bing = resultados[1][1]
    print()
    if azure and bing:
        print('      A rede esta bem. Se o mapa continua branco, o problema e a')
        print('      definicao do visual Azure Maps - so o administrador do')
        print('      Power BI da organizacao a pode ligar.')
        return 0
    if not azure and bing:
        print('      O Bing responde mas o Azure Maps NAO: e a rede/proxy a')
        print('      bloquear atlas.microsoft.com. Pedir para libertar esse')
        print('      endereco, ou trocar o visual pelo mapa normal do Power BI.')
        return 1
    print('      Nenhum servidor de mapas responde: este computador nao tem')
    print('      acesso a internet a partir do Python/Power BI (proxy?).')
    return 1


if __name__ == '__main__':
    raise SystemExit(main())
