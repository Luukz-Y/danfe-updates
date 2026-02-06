#!/usr/bin/env python3
"""
Sistema de Gerenciamento de Versões - DANFE Compactador
"""

# Versão atual do aplicativo
CURRENT_VERSION = "1.0.1"

# URL base para verificação de atualizações (será configurada quando hospedar)
UPDATE_BASE_URL = "https://Luukz-Y.github.io/danfe-updates"

# Informações da versão
VERSION_INFO = {
    "version": CURRENT_VERSION,
    "build": "2024.01.15",
    "release_date": "2024-01-15",
    "changelog": [
        "Versão inicial do DANFE Compactador",
        "Compactação de 3 DANFEs por folha",
        "Índices automáticos [x/y]",
        "Destaque de códigos SKU",
        "Sistema de licenciamento"
    ],
    "features": [
        "Compactação inteligente de PDFs",
        "Detecção automática de SKUs",
        "Reordenação por grupos de fração",
        "Interface moderna com CustomTkinter",
        "Abertura automática para impressão"
    ]
}

def get_version():
    """Retorna a versão atual"""
    return CURRENT_VERSION

def get_version_info():
    """Retorna informações completas da versão"""
    return VERSION_INFO

def compare_versions(version1, version2):
    """
    Compara duas versões no formato semver (x.y.z)
    Retorna: -1 se version1 < version2, 0 se iguais, 1 se version1 > version2
    """
    def version_tuple(v):
        return tuple(map(int, v.split('.')))
    
    v1_tuple = version_tuple(version1)
    v2_tuple = version_tuple(version2)
    
    if v1_tuple < v2_tuple:
        return -1
    elif v1_tuple > v2_tuple:
        return 1
    else:
        return 0

def is_newer_version(remote_version):
    """Verifica se a versão remota é mais nova que a atual"""
    return compare_versions(CURRENT_VERSION, remote_version) < 0
