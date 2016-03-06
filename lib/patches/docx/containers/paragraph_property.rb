module Docx
  module Elements
    module Containers
      class ParagraphProperty
        include ::Docx::Elements::Containers::Container
        include ::Docx::Elements::Element

        def self.tag
          'pPr'
        end

        def initialize(node)
          @node = node
        end

        def pstyle_node
          @node.children.find { |child| child.name == "pStyle" }
        end

        def pstyle_value
          pstyle_node.attributes["val"].value
        end
      end
    end
  end
end
