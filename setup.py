from setuptools import setup, find_packages

setup(
    name="Scripts_Alexo",  # Nome do pacote
    version="1.0.0",       # Versão inicial
    packages=find_packages(),  # Detecta automaticamente os subpacotes
    install_requires=[],      # Lista de dependências, se houver
    description="Scripts para gerenciamento de bots",
    author="Seu Nome",
    author_email="seuemail@exemplo.com",
    url="https://github.com/seuusuario/Scripts_Alexo",  # Repositório (opcional)
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",  # Versão mínima do Python
)
