module Pandocxtyle
  module Patches
    refine ::Docx::Elements::Containers::TextRun do
      def append_fld_char(type)
        fld_char = Nokogiri::XML::Node.new("w:fldChar", @node)
        fld_char["w:fldCharType"] = type
        @node.add_child(fld_char)
        self
      end

      def append_instr_text(content, attributes = {})
        instr_text = Nokogiri::XML::Node.new("w:instrText", @node)
        attributes.each { |key, value| instr_text[key] = value }
        instr_text.content = content
        @node.add_child(instr_text)
        self
      end

      def append_property(*nodes)
        prop = Nokogiri::XML::Node.new("w:rPr", @node)
        nodes.each { |n| prop.add_child(n) }
        @node.add_child(prop)
        self
      end

      def append_no_proof
        no_proof = Nokogiri::XML::Node.new("w:noProof", @node)
        append_property(no_proof)
      end
    end
  end
end
