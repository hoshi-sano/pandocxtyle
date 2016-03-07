module Pandocxtyle
  module Patches
    refine ::Docx::Elements::Containers::Table do
      def style
        if @node.xpath("w:tblPr/w:tblStyle/@w:val").first
          @node.xpath("w:tblPr/w:tblStyle/@w:val").first.value
        else
          ""
        end
      end

      def style=(new_style)
        if @node.xpath("w:tblPr/w:tblStyle/@w:val").first
          @node.xpath("w:tblPr/w:tblStyle/@w:val").first.value = new_style
        else
          prop = @node.xpath("w:tblPr").first
          style = Nokogiri::XML::Node::new("w:tblStyle", prop)
          style["w:val"] = new_style
          prop.add_child(style)
        end
      end
    end
  end
end
