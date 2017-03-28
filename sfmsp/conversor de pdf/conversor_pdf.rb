require 'rubygems'
require 'pdf-reader'
require 'writeexcel'

caminho = File.expand_path File.dirname(__FILE__)
Dir.glob(caminho + "/*.pdf").each do |arquivo|

  reader = PDF::Reader.new(arquivo)

  def extrai_tabela(pagina)
    tabela = ''
    pagina.each_line do |line|
      if line.to_s.start_with?('     0')
        tabela = tabela + line
      end
    end
    tabela
  end

  def eh_num?(a)
    if a == '0'
      return true
    elsif (a != 0) && a.to_i == 0
      return false
    else
      return true
    end
  end

  def gera_excel(array, arquivo)
    workbook = WriteExcel.new(arquivo.chomp('.pdf') + '.xls')
    worksheet  = workbook.add_worksheet
    i = 1

    worksheet.write(0, 0, 'Cracha')
    worksheet.write(0, 1, 'Nome')
    worksheet.write(0, 2, 'Emp')
    worksheet.write(0, 3, 'Depto')
    worksheet.write(0, 4, 'Setor')
    worksheet.write(0, 5, 'Seção')
    worksheet.write(0, 6, 'Centro Custo')
    worksheet.write(0, 7, 'Registro')
    worksheet.write(0, 8, 'Horario')
    worksheet.write(0, 9, 'Limite')


    array.each do |tabela|
      tabela.each_line do |linha|
          #Cracha tem 10 digitos, o nome eh indefinido, depto tem 5, setor tem 5, secao tem 5, centro tem 4 digitos, um traco e mais 6 digitos, registro tem 6, horario tem 4 e limite tem ate 3, dois pontos e dois digitos

          #pegando os 10 digitos do cracha
          worksheet.write(i, 0, linha[5, 10])
          pos_linha = 16 # posicao do vetor linha sempre sera essa apos o cracha
          ant_linha = 0

          #pegando o nome
          nome = ''
          while !eh_num?(linha[pos_linha])
            nome = nome + linha[pos_linha]
            pos_linha = pos_linha + 1
          end
          worksheet.write(i, 1, nome)

          #pegando o emp
          ant_linha = pos_linha
          tamanho = 0
          while eh_num?(linha[pos_linha])
            tamanho = tamanho + 1
            pos_linha = pos_linha + 1
          end
          worksheet.write(i, 2, linha[ant_linha, tamanho])

          #andando o espaço em branco
          while !eh_num?(linha[pos_linha])
            pos_linha = pos_linha + 1
          end

          #pegando o depto
          ant_linha = pos_linha
          tamanho = 0
          while eh_num?(linha[pos_linha])
            tamanho = tamanho + 1
            pos_linha = pos_linha + 1
          end
          worksheet.write(i, 3, linha[ant_linha, tamanho])

          #andando o espaço em branco
          while !eh_num?(linha[pos_linha])
            pos_linha = pos_linha + 1
          end

          #pegando o setor
          ant_linha = pos_linha
          tamanho = 0
          while eh_num?(linha[pos_linha])
            tamanho = tamanho + 1
            pos_linha = pos_linha + 1
          end
          worksheet.write(i, 4, linha[ant_linha, tamanho])

          #andando o espaço em branco
          while !eh_num?(linha[pos_linha])
            pos_linha = pos_linha + 1
          end

          #pegando a secao
          ant_linha = pos_linha
          tamanho = 0
          while eh_num?(linha[pos_linha])
            tamanho = tamanho + 1
            pos_linha = pos_linha + 1
          end
          worksheet.write(i, 5, linha[ant_linha, tamanho])

          #andando o espaço em branco
          while !eh_num?(linha[pos_linha])
            pos_linha = pos_linha + 1
          end

          #pegando o centro custo
          ant_linha = pos_linha
          tamanho = 0
          while eh_num?(linha[pos_linha])
            tamanho = tamanho + 1
            pos_linha = pos_linha + 1
          end
          pos_linha = pos_linha + 1 # pulando o caracter especial -
          while eh_num?(linha[pos_linha])
            tamanho = tamanho + 1
            pos_linha = pos_linha + 1
          end
          worksheet.write(i, 6, linha[ant_linha, tamanho])

          #andando o espaço em branco
          while !eh_num?(linha[pos_linha])
            pos_linha = pos_linha + 1
          end

          #pegando o registro
          ant_linha = pos_linha
          tamanho = 0
          while eh_num?(linha[pos_linha])
            tamanho = tamanho + 1
            pos_linha = pos_linha + 1
          end
          worksheet.write(i, 7, linha[ant_linha, tamanho])

          #andando o espaço em branco e verificando se existe horario ou vai direto pro limite
          count_horario = 0
          while !eh_num?(linha[pos_linha])
            pos_linha = pos_linha + 1
            count_horario = count_horario + 1
          end


          if count_horario > 8 #se não houver horas, ir direto pro limite

            #pegando o limite
            ant_linha = pos_linha
            tamanho = 0
            while eh_num?(linha[pos_linha])
              tamanho = tamanho + 1
              pos_linha = pos_linha + 1
            end
            pos_linha = pos_linha + 1 # pulando o caracter especial :
            while eh_num?(linha[pos_linha])
              tamanho = tamanho + 1
              pos_linha = pos_linha + 1
            end
            worksheet.write(i, 9, linha[ant_linha, tamanho+1])

          else

            #pegando o horario
            ant_linha = pos_linha
            tamanho = 0
            while eh_num?(linha[pos_linha])
              tamanho = tamanho + 1
              pos_linha = pos_linha + 1
            end
            worksheet.write(i, 8, linha[ant_linha, tamanho])

            #andando o espaço em branco
            while !eh_num?(linha[pos_linha])
              pos_linha = pos_linha + 1
            end

            #pegando o limite
            ant_linha = pos_linha
            tamanho = 0
            while eh_num?(linha[pos_linha])
              tamanho = tamanho + 1
              pos_linha = pos_linha + 1
            end
            pos_linha = pos_linha + 1 # pulando o caracter especial :
            while eh_num?(linha[pos_linha])
              tamanho = tamanho + 1
              pos_linha = pos_linha + 1
            end
            worksheet.write(i, 9, linha[ant_linha, tamanho+1])
          end

          #desce para a proxima linha do excel
          i = i + 1
      end
    end
    workbook.close
  end

  #-------Main-------
  array = []
  reader.pages.each do |page|
    array << extrai_tabela(page.text)
  end

  gera_excel(array, arquivo)
end
