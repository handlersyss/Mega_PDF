# Mega PDF

Este repositório contém um script robusto que permite:

Junção de arquivos PDF
Conversão entre formatos Word e PDF (e vice-versa)
Conversão entre PDF e Excel (e vice-versa)
Impressão de arquivos sem limites

O script é compatível tanto com Linux quanto com Windows. Para utilizar todas as funcionalidades, certifique-se de ter todos os requisitos instalados a partir dos arquivos .txt fornecidos.

Para realizar impressões, é necessário ter um leitor de PDF instalado e associado ao tipo de arquivo PDF no Windows. No Linux, certifique-se de que o CUPS (Common UNIX Printing System) esteja instalado e funcionando, pois o comando lp depende dele.

Este programa é ideal para quem precisa gerenciar e converter documentos de maneira eficiente e sem restrições.

OBS: Para evitar erro na hora de fazer a conversão do pdf para excel por falta da biblioteca JVM(Java Virtual Machine) necessária para a execução do Tabula, que é uma ferramenta Java.

-----

Aqui estão os passos que você pode seguir para resolver esse problema:

Instale o Java: Você precisa ter o Java Development Kit (JDK) instalado em seu sistema. Você pode baixá-lo do site oficial do Oracle ou usar o OpenJDK.

Defina a variável de ambiente JAVA_HOME: Após instalar o Java, você precisa definir a variável de ambiente JAVA_HOME para apontar para o diretório de instalação do JDK. Aqui está como fazer isso no Windows:

Clique com o botão direito no ícone "Meu Computador" ou "Este PC" na área de trabalho ou no explorador de arquivos e selecione "Propriedades".
Clique em "Configurações avançadas do sistema".
Clique em "Variáveis de ambiente".
Em "Variáveis do sistema", clique em "Novo".
Adicione JAVA_HOME como o nome da variável e o caminho do diretório do JDK (por exemplo, C:\Program Files\Java\jdk-11.0.2) como o valor da variável.
Clique em "OK" para salvar as mudanças.
Verifique se o Java está funcionando: Abra o Prompt de Comando e digite java -version para verificar se o Java está corretamente instalado e configurado.

Reinicie o Python: Certifique-se de reiniciar qualquer script ou terminal Python para que ele reconheça as novas variáveis de ambiente.

----

![image](https://github.com/handlersyss/Mega_PDF/assets/169811777/62fb8aca-f8d3-4d16-aa33-9637874cb015)

----

Feedback e sugestões são bem-vindos! Sinta-se à vontade para compartilhar suas opiniões, ideias e melhorias para este script. Sua contribuição é muito apreciada.