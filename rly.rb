#!/bin/env ruby
#==============================================================================
#
# rly.rb converts an RLY markup document into Latex, PDF, HTML, or Word.
#
# Usage: rly.rb <doc|pdf|tex|htm> <file.rly>
#
#==============================================================================

#==============================================================================
#
# Reader module for RLY format
#
#==============================================================================

module Reader

  class Text
    attr_accessor :text
  end

  class Header < Text
    attr_accessor :title, :author, :date, :doctype
    def initialize(rly_text)
      @text = rly_text
      @text.split(/\n/).each do |elem|
        case elem
        when /^\\title\s+(.*)/
          @title = $1.strip
          @title.sub!(/\[(\w+)\]\z/, "")
          @doctype = $1
        when /^\\author\s+(.*)/
          @author = $1.strip
        when /^\\company\s+(.*)/
          @author = $1.strip
        when /^\\date\s+(.*)/
          @date = $1.strip
        else
          $stderr.puts "can't parse #{elem}"
          exit(1)
        end
      end
    end
  end

  class Heading < Text
    attr_accessor :title, :level, :ref
    def initialize(rly_text)
      @text = rly_text
      @text =~ /^\\h(\d+)\s+(.*)/
      @level = $1.to_i
      @title = $2.sub(/\s+\[(\w+)\]/,"").strip
      @ref   = $1
    end
  end

  class Paragraph < Text
    def initialize(rly_text)
      @text = rly_text.gsub(/\n/," ")
    end
  end

  class Verbatim < Text
    attr_accessor :level
    def initialize(rly_text)
      @text = rly_text
      @text =~ /^(\s+)/
      @level = $1.length
    end
  end

  class Latex < Text
    def initialize(rly_text)
      @text = rly_text
    end
  end

  class Enumeration < Text
    attr_accessor :entries
    def initialize(rly_text)
      @text = rly_text
      @entries = @text.split(/^\d+\. /)[1..-1].map { |e| e.gsub!(/\n/," ") }
      @entries.map! { |e| e.gsub(/\s{2,}/," ") }
    end
  end

  class Bullet < Enumeration
    def initialize(rly_text)
      @text = rly_text
      @entries = @text.split(/^\* /)[1..-1].map { |e| e.gsub!(/\n/," ") }
      @entries.map! { |e| e.gsub(/\s{2,}/," ") }
    end
  end

  class Table < Enumeration
    attr_accessor :caption, :format, :ref, :option
    def initialize(rly_text)
      @text = rly_text
      rows = @text.split(/\n/)
      #rows.pop
      ttop = rows.shift
      tokens = ttop.split
      tokens.shift
      @option  = tokens.pop
      @format  = @option =~ /\[/ ? tokens.pop : @option
      @ref     = tokens.pop
      @caption = tokens.join(" ")[1..-2]
      @entries = rows
      @entries.map! { |e| e.split(/\&/).map { |i| i.strip } }
      ncols = @entries.first.size
      @entries.map! { |e| e + Array.new(ncols - e.size, "") }
    end
  end

  class Tabular < Enumeration
    def initialize(rly_text)
      @text = rly_text
      rows = @text.split(/\n/)
      ttop = rows.shift
      tokens = ttop.split
      tokens.shift
      @entries = rows
      @entries.map! { |e| e.split(/\&/).map { |i| i.strip } }
      ncols = @entries.first.size
      @entries.map! { |e| e + Array.new(ncols - e.size, "") }
    end
  end

  class Table2 < Enumeration
    attr_accessor :caption, :format, :ref, :option
    def initialize(rly_text)
      @text = rly_text
      rows = @text.split(/\n/)
      tokens = rows.shift.split
      @option  = ""
      @format  = tokens.shift
      @ref     = tokens.pop
      @caption = tokens.join(" ").gsub(/\"/,"")
      @entries = rows.map { |e| e.split(/\:/).map { |i| i.strip } }
      ncols = @entries.first.size
      @entries.map! { |e| e + Array.new(ncols - e.size, "") }
    end
  end

  class Insert < Text
    attr_accessor :file, :path, :option, :caption, :ref
    def initialize(rly_text)
      @text    = rly_text
      tokens   = @text.split[1..-1]
      @file    = tokens.shift
      @path    = @file.sub(/\.fig\z/,"")
      @option  = tokens.pop if tokens.last =~ /\[/
      @ref     = tokens.pop if tokens.last =~ /^\w+\z/
      @caption = tokens.join(" ")[1..-2]
    end
  end

  class Section < Enumeration
    attr_accessor :heading
    def initialize()
      @entries = Array.new
    end
  end

  class Document
    attr_reader :header, :sections, :document
    def initialize(rly_file)
      @sections = Array.new
      @document = IO.readlines(recurse_insert_rly(rly_file)).map do |line|
        line =~ /^%|\\(comment|end)/ ? "\n" : line
      end
      @document.join.split(/^\n/).each do |rly_text|
        next if rly_text.empty?
        case rly_text
        when /^\\comment(.*)/,/^\\end/
        when /^\\title/,/^\\author/,/^\\company/,/^\\date/
          @header = Header.new(rly_text)
        when /^\\insert/
          @sections.last.entries.push Insert.new(rly_text)
        when /^\\h(\d+)/
          @sections.push Section.new
          @sections.last.heading = Heading.new(rly_text)
        when /^\d+\. /
          @sections.last.entries.push Enumeration.new(rly_text)
        when /^\* /
          @sections.last.entries.push Bullet.new(rly_text)
        when /^\\table/
          @sections.last.entries.push Table.new(rly_text)
        when /^\\tabular/
          @sections.last.entries.push Tabular.new(rly_text)
        when /^\|(c|p)/
          @sections.last.entries.push Table2.new(rly_text)
        when /^\s+/
          @sections.last.entries.push Verbatim.new(rly_text)
        when /^\\?\w+/, /^\$/
          @sections.last.entries.push Paragraph.new(rly_text)
        else
          $stderr.puts "Error: can't parse\n" + rly_text
        end
      end
    end
    private
    def recurse_insert_rly(rly_file)
      inlined_rly = Array.new
      IO.readlines(rly_file).each do |line|
        line = IO.read($1) if line =~ /^\\insert\s+(.*\.rly)/
        inlined_rly.push line
      end
      File.open(rly_file+".rly","w") do |fd|
        fd.puts inlined_rly
      end
      return rly_file+".rly"
    end
  end

end # module Reader

#==============================================================================
#
# Writer module to convert RLY format to PDF, MSWord, LaTex, HTML
#
#==============================================================================

module Writer

  require "pathname"

  #------------------------------------------------------------------------------
  # TexWriter converts .rly -> .tex (Latex)

  class TexWriter
    def initialize(rly_file)
      @rly_file = rly_file
      tex_file = rly_file.sub(/rly\z/,"tex")
      doc = Document.new(rly_file)
      tex = Array.new
      tex << header(doc.header) + "\n"
      doc.sections.each do |sec|
        tex << heading(sec.heading) + "\n"
        sec.entries.each do |blk|
          case blk
          when Paragraph
            tex << paragraph(blk) + "\n"
          when Insert
            tex << insert(blk) + "\n"
          when Table, Table2
            tex << table(blk) + "\n"
          when Tabular
            tex << tabular(blk) + "\n"
          when Bullet
            tex << bullet(blk) + "\n"
          when Enumeration
            tex << enumeration(blk) + "\n"
          when Verbatim
            tex << verbatim(blk) + "\n"
          end
        end
      end
      tex << '\end{document}'
      File.open(tex_file,"w") do |fd|
        fd.puts tex
      end
      system "rm -f *.rly.rly"
    end
    def inlines(txt)
      txt.gsub!(/\bP\{/, "page~\\pageref{")    # P{pag_ref}
      txt.gsub!(/\bf\{/, "Figure~\\ref{")      # f{fig_ref}
      txt.gsub!(/\bt\{/, "Table~\\ref{")       # t{tab_ref}
      txt.gsub!(/\bs\{/, "Section~\\ref{")     # s{sec_ref}
      txt.gsub!(/\bS\{/, "\\nameref{")         # S{sec_ref}
      txt.gsub!(/\be\{/, "\\emph{")            # e{emphasized}
      txt.gsub!(/\bi\{/, "\\textit{")          # i{italic}
      txt.gsub!(/\bb\{/, "\\textbf{")          # b{bold}
      txt.gsub!(/\bu\{/, "\\underline{")       # u{underline}
      txt.gsub!(/\bc\{/, "\\texttt{")          # c{code}
      txt.gsub!(/\bF\{/, "\\footnote{")        # F{footnote}
      if txt =~ /\bv\{/                        # v{verbatim}
        txt.gsub!(/\bv\{/, "\\verb+")
        txt.gsub!(/\}/, "+")
      end
      if txt =~ /\bm\{(.*)\}/
        txt.gsub!(/m\{(.*)\}/, '$' + $1 + '$') # m{math}
      end
      return txt
    end
    def escapes(txt)
      txt.gsub!(/\#/, "\\#")                   # escape #
      txt.gsub!(/\%/, "\\%")                   # escape %
      txt.gsub!(/\_/, "\\_")                   # escape _
      txt.gsub!(/\~/, "\\~")                   # escape ~
      return txt
    end
    def header(header)
      return ""+
        "\\documentclass[12pt,letterpaper]{article}\n"+
        "\\usepackage{graphicx,longtable,float}\n"+
        "\\usepackage[margin=0.5in,bottom=1in]{geometry}\n"+
        "\\title{#{header.title}}\n"+
        "\\author{#{header.author}}\n"+
        "\\date{#{header.date}}\n"+
        "\\begin{document}\n"+
        "\\pagestyle{plain}\n"+
        "\\pagenumbering{arabic}\n"+
        "\\maketitle"
    end
    def heading(hd)
      sub   = "sub" * (hd.level-1)
      label = "\n\\label{#{hd.ref}}" if !hd.ref.nil?
      return "\\#{sub}section{#{escapes(hd.title)}}#{label}\n"
    end
    def wrap(s, width=72)
      s.gsub(/(.{1,#{width}})(\s+|\Z)/, "\\1\n")
    end
    def paragraph(pp)
      return wrap( inlines(escapes(pp.text)) + "\n" )
    end
    def verbatim(enum)
      return ""+
        "\\begin{verbatim}\n" +
        enum.text +
        "\\end{verbatim}\n"
    end
    def latex(txt)
      return txt
    end
    def enumeration(enum)
      return ""+
        "\\begin{enumerate}\n" +
        enum.entries.map{|e|"\\item #{escapes(inlines(e))}\n"}.join +
        "\\end{enumerate}\n"
    end
    def bullet(enum)
      return ""+
        "\\begin{list}{$\\bullet$}{\n"+
        "\\setlength{\\topsep}{0.1ex}\n"+
        "\\setlength{\\parsep}{0.1ex}}\n"+
        enum.entries.map{|e|"\\item #{escapes(inlines(e))}\n"}.join +
        "\\end{list}"
    end
    def table(tbl)
      label = "\\label{#{tbl.ref}}\n" if !tbl.ref.nil?
      items = tbl.entries.map { |e| escapes(inlines(e.join(' & '))) }
      heads = items.first
      return ""+
        "\\begin{longtable}{#{tbl.format}}\n"+
        "\\caption{#{escapes(tbl.caption)}}#{label.chomp}\\\\\n"+
        "\\hline\n"+heads+"\\\\\n\\hline\\hline\n"+
        "\\endfirsthead\\caption[]{(continued)}\\\\\n\\hline\n"+
        heads+"\\\\\n\\hline\\hline\n"+
        "\\endhead\n"+
        items[1..-1].flatten.join("\\\\\n\\hline\n")+
        "\\\\\n\\hline\n"+
        "\\end{longtable}\n"
    end
    def tabular(tbl)
      items = tbl.entries.map { |e| escapes(inlines(e.join(' & '))) }
      ncols = tbl.entries.first.size
      return ""+
        "\\bigskip\n\n"+
        "\\begin{tabular}{"+ ("l" * ncols) + "}\n"+
        items.join("\\\\\n")+
        "\\\\\n\\end{tabular}\n"+
        "\n\\bigskip\n"
    end
    def insert(ins)
      system "fig2dev -L eps #{ins.file} #{ins.path}.eps" if !File.exist?(ins.path+".eps")
      bounding_box = `head -6 #{ins.path}.eps`.split(/\n/).last.split[-2].to_i
      ins.ref = File.basename(ins.path).split(/_/).map{|e|e.capitalize}.join("")
      graphics_option = bounding_box > 500 ? "[width=\\textwidth]" : ""
      return ""+
        "\\begin{figure}[H]\n"+
        "  \\center\n"+
        "  \\includegraphics#{graphics_option}{#{ins.path}.eps}\n"+
        "  \\caption{#{escapes(ins.caption)}}\n"+
        "  \\label{#{ins.ref}}\n"+
        "\\end{figure}\n"
    end
    def section(sec)
      return ""+
        sec.heading.latex() +
        "\n\n"+
        sec.entries.map { |e| e.latex() }.join("\n\n")
    end
  end

  #------------------------------------------------------------------------------
  # PdfWriter converts .rly -> .tex (see TexWriter) -> .pdf (Adobe
  # Portable Document Format)

  class PdfWriter
    def initialize(rly_file)
      TexWriter.new(rly_file)
      doc_name = rly_file.sub(/\.rly\z/,"")
      system "latex  #{doc_name}.tex"
      system "latex  #{doc_name}.tex"
      system "dvipdf #{doc_name}.dvi"
      tmp_files = %w(aux dvi log out toc tex rly.rly).map{|x| doc_name+"."+x }
      system "rm -f "+tmp_files.join(" ")
    end
  end

  #------------------------------------------------------------------------------
  # HtmWriter converts .rly -> .htm (Hypertext Markup Language)

  class HtmWriter
    def initialize(rly_file)
      htm_file = rly_file.sub(/rly\z/,"htm")
      doc = Document.new(rly_file)
      htm = Array.new
      htm << '<!DOCTYPE html>'
      htm << '<html>'
      htm << '<head>'
      htm << IO.read(File.dirname(__FILE__)+"/yourstyle.css")
      htm << "<title>#{doc.header.title}</title>"
      htm << '</head>'
      htm << '<body>'
      htm << '<div id="container">'
      htm << '<div id="content">'
      doc.sections.each do |sec|
        htm << heading(sec.heading)
        sec.entries.each do |blk|
          case blk
          when Paragraph
            htm << paragraph(blk) + "\n"
          when Insert
            htm << insert(blk) + "\n"
          when Table, Table2
            htm << table(blk) + "\n"
          when Bullet
            htm << bullet(blk) + "\n"
          when Enumeration
            htm << enumeration(blk) + "\n"
          when Verbatim
            htm << verbatim(blk) + "\n"
          end
        end
      end
      htm << '</div>'
      htm << '<div id="footer">'
      htm << 'Acme Corporation'
      htm << '</div>'
      htm << '</div>'
      htm << '</body>'
      htm << '</html>'
      File.open(htm_file,"w") do |fd|
        fd.puts htm
      end
    end
    private
    def heading(hd)
      aref = "id=\"#{hd.ref}\"" if !hd.ref.nil?
      head = "<h#{(hd.level)}#{aref}>#{hd.title}</h#{(hd.level)}>"
      return head + "\n"
    end
    def paragraph(pp)
      return "<p>#{pp.text}</p>\n"
    end
    def verbatim(enum)
      return "<pre>#{enum.text}</pre>\n"
    end
    def enumeration(enum)
      return "<ol>\n"+
        enum.entries.map{|e|"<li>#{e}</li>\n"}.join +
        "</ol>\n"
    end
    def bullet(enum)
      return "<ul>\n"+
        enum.entries.map{|e|"<li>#{e}</li>\n"}.join +
        "</ul>\n"
    end
    def table(tbl)
      return ""
    end
    def insert(ins)
      system "fig2dev -L png #{ins.file} #{ins.path}.png" if !File.exist?(ins.path+".png")
      return "<img src=\"#{ins.path}.png\" id=\"#{ins.ref}\" alt=\"#{ins.caption}\"\n"
    end
    def section(sec)
      return ""
    end
  end

  #------------------------------------------------------------------------------
  # DocWriter converts .rly -> .doc (Microsoft Word)

  class DocWriter
    def initialize(rly_file)
      require "win32ole"
      @xref_figures = %w(nada)
      @xref_tables  = %w(nada)
      @const = {
        :wdAlignParagraphJustify => 3,
        :wdAlignPageNumberCenter => 1,
        :wdTableFormatColorful2  => 9,
        :wdStory                 => 6,
        :wdFindContinue          => 1,
        :wdReplaceAll            => 2,
        :wdHeaderFooterPrimary   => 1,
        :wdCaptionPositionBelow  => 1,
        :wdCaptionPositionAbove  => 0,
        :wdFindContinue          => 1,
        :wdReplaceAll            => 2
      }
      dir = File.dirname(__FILE__)
      rly = Document.new(rly_file)
      wrd = WIN32OLE.new('Word.Application')
      wrd.Visible = true if false
      wrd.Documents.Add(false ? "Normal" : dir+"/yourstyle.dot")
      doc = wrd.ActiveDocument
      sel = wrd.Selection
      rly.sections.each do |sec|
        sel.Style = "Heading #{sec.heading.level}"
        sel.TypeText sec.heading.title
        sel.TypeParagraph
        sec.entries.each do |blk|
          sel.Style = "Normal"
          sel.ParagraphFormat.Alignment = @const[:wdAlignParagraphJustify]
          case blk
          when Paragraph
            sel.TypeParagraph
            sel.TypeText blk.text
            sel.TypeText "\n"
          when Insert
            sel.TypeParagraph
            file = "#{Dir.pwd}/#{blk.file}".sub(/\.fig\z/,".eps")
            # todo: how should we map paths from unix to windows?
            file.gsub!(/\//, "\\")
            pic = sel.InlineShapes.AddPicture({
                                                'FileName'         => file,
                                                'LinkToFile'       => false,
                                                'SaveWithDocument' => true
                                              })
            if file =~ /cgm\z/
              pic.ScaleHeight = 150
              pic.ScaleWidth  = 150
            end
            sel.TypeText "\n"
            sel.TypeParagraph
            sel.InsertCaption({
                                'Label'     => "Figure",
                                'Title'     => ": #{blk.caption}",
                                'Position'  => @const[:wdCaptionPositionBelow]
                              })
            sel.ParagraphFormat.Alignment = @const[:wdAlignPageNumberCenter]
            if !blk.ref.nil?
              doc.Bookmarks.Add({'Range' => sel.Range, 'Name' => blk.ref})
            end
            @xref_figures.push blk.ref
            sel.TypeText "\n"
          when Table, Table2
            sel.TypeParagraph
            rows = blk.entries.size
            cols = blk.entries.first.size
            tbl = doc.Tables.Add({
                                   'Range'      => sel.Range,
                                   'NumRows'    => rows,
                                   'NumColumns' => cols
                                 })
            rows.times do |r|
              cols.times do |c|
                tbl.Cell(r+1,c+1).Range.Text = blk.entries[r][c]
              end
            end
            sel.InsertCaption({
                                'Label'     => "Table",
                                'Title'     => ": #{blk.caption}",
                                'Position'  => @const[:wdCaptionPositionAbove]
                              })
            sel.ParagraphFormat.Alignment = @const[:wdAlignPageNumberCenter]
            if !blk.ref.nil?
              doc.Bookmarks.Add({'Range' => sel.Range, 'Name' => blk.ref})
            end
            @xref_tables.push blk.ref
            wdTableFormatList4 = 27
            wdTableFormatGrid8 = 23
            wdTableFormatGrid4 = 19
            wdTableFormatGrid3 = 18
            wdTableFormatGrid1 = 16
            tbl.AutoFormat({
                             'Format'       => wdTableFormatList4,
                             'ApplyBorders' => true,
                             'ApplyFont'    => true,
                             'ApplyColor'   => true
                           })
            sel.Move({'Unit'=>@const[:wdStory]})
          when Bullet
            sel.TypeParagraph
            sel.Range.ListFormat.ApplyBulletDefault
            blk.entries.each do |e|
              sel.TypeText e
              sel.TypeParagraph
            end
          when Enumeration
            sel.TypeParagraph
            sel.Range.ListFormat.ApplyNumberDefault
            blk.entries.each do |e|
              sel.TypeText e
              sel.TypeParagraph
            end
          when Verbatim
            sel.TypeParagraph
            sel.Font.Name = "Courier New"
            sel.Font.Size = 8
            sel.TypeText blk.text
          end
        end
      end
      inlines(sel, "b", "Bold")
      inlines(sel, "i", "Italic")
      inlines(sel, "e", "Italic")
      inlines(sel, "u", "Underline")
      replace(sel, "s", "Section")
      hyperlink(doc, sel, "t")
      hyperlink(doc, sel, "f")
      modify_header(doc, "DocumentName", rly.header.title)
      sel.HomeKey(unit=6)
      doc.SaveAs(Dir.pwd+"/"+rly_file.sub(/rly\z/,"doc"))
      doc.Close()
      wrd.Quit()
    end
    private
    def inlines(sel, mark, format)
      sel.HomeKey(unit=6)
      sel.Find.MatchWildcards = true
      sel.Find.Text = "#{mark}[{]*[}]"
      while sel.Find.Execute
        sel.Text =~ /\{(.*?)\}/
        eval "sel.Font.#{format} = true"
        sel.TypeText $1
      end
    end
    def replace(sel, mark, text, for_rly=false)
      sel.HomeKey(unit=6)
      sel.Find.MatchWildcards = true
      if for_rly
        sel.Find.Text = "#{mark}*"
      else
        sel.Find.Text = "#{mark}[{]*[}]"
      end
      while sel.Find.Execute
        sel.TypeText text
      end
    end
    def bookmark(doc, sel, name)
      doc.Bookmarks.Add({'Range'=>sel.Range, 'Name'=>name})
    end
    def hyperlink(doc, sel, mark)
      sel.HomeKey(unit=6)
      sel.Find.MatchWildcards = true
      sel.Find.Text = "#{mark}[{]*[}]"
      while sel.Find.Execute
        sel.Text =~ /\{(.*?)\}/
        ref = $1
        cap = case mark
              when "f" then "Figure #{@xref_figures.index(ref)}"
              when "t" then "Table #{@xref_tables.index(ref)}"
              end
        doc.Hyperlinks.Add({
                             'Anchor'        => sel.Range,
                             'SubAddress'    => ref,
                             'TextToDisplay' => cap
                           })
      end
    end
    def modify_header(doc, mark, text)
      doc.Sections(1).Headers(1).Range.StoryType
      doc.StoryRanges.each do |rng|
        rng.Find.Text = mark
        rng.Find.Replacement.Text = text
        rng.Find.Wrap = @const[:wdFindContinue]
        rng.Find.Execute({'Replace' => @const[:wdReplaceAll]})
      end
    end
  end

end # module Writer

#==============================================================================
# MAIN
#==============================================================================
include Reader
include Writer

case (output_format = ARGV.shift)
when "doc", "pdf", "tex", "htm"
  eval output_format.capitalize + "Writer.new(*ARGV)"
else
  $stderr.puts "Error: format '#{output_format}' not supported\n" +
    "Usage: rly.rb <doc|pdf|tex|htm> <file.rly>"
  exit(1)
end
