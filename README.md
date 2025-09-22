# DANFE Compactador - Sistema de Atualizações

Este repositório contém os arquivos de atualização para o DANFE Compactador.

## 📋 Estrutura

- `version.json` - Informações da versão mais recente
- `updates/` - Diretório com arquivos de atualização
- `index.html` - Página de informações sobre atualizações

## 🔄 Como funciona

O aplicativo DANFE Compactador verifica automaticamente este repositório para atualizações disponíveis.

## 📦 Arquivos de Atualização

Cada atualização é um arquivo ZIP contendo:
- `index.py` - Aplicativo principal
- `version.py` - Informações de versão
- `updater.py` - Sistema de atualização
- `gerar_licencas.py` - Gerador de licenças
- `requirements.txt` - Dependências

## 🚀 Para Desenvolvedores

Para criar uma nova atualização:

1. Execute `python setup_github_updates.py`
2. Siga as instruções na tela
3. Faça commit e push dos arquivos gerados
4. O GitHub Pages atualizará automaticamente

## 📞 Suporte

Para suporte técnico, entre em contato com o desenvolvedor.
