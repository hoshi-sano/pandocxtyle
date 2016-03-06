module Pandocxtyle
  using Patches

  module Table
    DEFAULT_TARGET_STYLE = "TableNormal"
    DEFAULT_CAPTION_LABEL = "Table."

    module_function

    def sub_style
      params = ARGV.getopts("i:o:",
                            "target:#{DEFAULT_TARGET_STYLE}", "with:")
      Docx::Document.open(params["i"]) do |doc|
        doc.tables.each do |table|
          if table.style == params["target"]
            table.style = params["with"]
          end
        end
        doc.save(params["o"])
      end
    end

    def seq_caption
      params = ARGV.getopts("i:o:", "label:#{DEFAULT_CAPTION_LABEL}")
      Docx::Document.open(params["i"]) do |doc|
        caption_idx = 0
        doc.paragraphs.each do |paragraph|
          next unless paragraph.pstyle_is_a?("TableCaption")
          caption_idx += 1
          instr_text_attr = {"xml:space" => "preserve"}
          [
            ->(r){ r.text = params["label"]; r },
            ->(r){ r.append_fld_char("begin") },
            ->(r){ r.append_instr_text(" ", instr_text_attr) },
            ->(r){ r.append_instr_text("SEQ ", instr_text_attr) },
            ->(r){ r.append_instr_text(params["label"], instr_text_attr) },
            ->(r){ r.append_instr_text(" \\* ARABIC", instr_text_attr) },
            ->(r){ r.append_fld_char("separate") },
            ->(r){ r.append_instr_text(" ", instr_text_attr) },
            ->(r){ r.append_no_proof; r.text = caption_idx; r },
            ->(r){ r.append_fld_char("end") },
          ].inject(paragraph.properties.last) do |prev_node, func_for_current|
            text_run = Docx::Elements::Containers::TextRun.create_with(paragraph)
            text_run = func_for_current.call(text_run)
            text_run.insert_after(prev_node)
          end
        end
        doc.save(params["o"])
      end
    end
  end
end
