# 📄 Mega PDF

Uma ferramenta poderosa para manipulação de arquivos PDF com múltiplas funcionalidades.

## ✨ Funcionalidades

- 📎 Junção de arquivos PDF
- 🔄 Conversão entre formatos Word e PDF (bidirecional)
- 📊 Conversão entre PDF e Excel (bidirecional)
- 🖨️ Impressão de arquivos sem limitações
- 📦 Compactação de arquivos

## 💻 Compatibilidade

O script é compatível com:
- Windows
- Linux

## 📋 Pré-requisitos

### Windows
- Leitor de PDF instalado e associado aos arquivos .pdf
- Java Development Kit (JDK) para conversão PDF-Excel

### Linux
- CUPS (Common UNIX Printing System) instalado
- Dependências listadas no arquivo requirements_linux.txt

## 🚀 Instalação

1. Clone o repositório:
```bash
git clone https://github.com/handlersyss/Mega_PDF.git
```

2. Acesse o diretório:
```bash
cd Mega_PDF
```

3. Instale as dependências:
```bash
# Para Linux
pip3 install -r requirements_linux.txt

# Para Windows
pip3 install -r requirements_Windows.txt
```

4. Execute o programa:
```bash
python3 MEGA_PDF.py
```

## ⚠️ Configuração do Java (Importante)

Para a funcionalidade de conversão PDF para Excel, é necessário configurar corretamente o JVM (Java Virtual Machine):

1. Instale o Java Development Kit (JDK)
   - Baixe do site oficial da Oracle ou use OpenJDK

2. Configure a variável JAVA_HOME:
   - Abra as Propriedades do Sistema
   - Acesse "Configurações avançadas do sistema"
   - Clique em "Variáveis de ambiente"
   - Em "Variáveis do sistema", adicione JAVA_HOME
   - Defina o caminho do JDK (exemplo: C:\Program Files\Java\jdk-11.0.2)

3. Verifique a instalação:
```bash
java -version
```

## 📸 Interface

![Interface do Mega PDF](https://github.com/user-attachments/assets/aec8199e-28e9-4c9b-bfd4-11016bd12a46)

## ⚠️ Nota sobre Linux

Algumas funcionalidades podem apresentar limitações no sistema Linux. Atualizações futuras trarão melhorias para total compatibilidade.

## 🤝 Contribuição

Feedback e sugestões são muito bem-vindos! Sinta-se à vontade para:
- Reportar bugs
- Sugerir novas funcionalidades
- Melhorar a documentação
- Enviar pull requests

---