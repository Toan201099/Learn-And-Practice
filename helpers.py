import aspose.words as aw
import helps
from helps import verify_parameter_nodes
from helps import get_ancestor_in_body
from helps import process_marker
def extract_content(startNode : aw.Node, endNode : aw.Node, isInclusive : bool):
    
    # First, check that the nodes passed to this method are valid for use.
    verify_parameter_nodes(startNode, endNode)

    # Create a list to store the extracted nodes.
    nodes = []

    # If either marker is part of a comment, including the comment itself, we need to move the pointer
    # forward to the Comment Node found after the CommentRangeEnd node.
    if (endNode.node_type == aw.NodeType.COMMENT_RANGE_END and isInclusive) :
        
        node = find_next_node(aw.NodeType.COMMENT, endNode.next_sibling)
        if (node != None) :
            endNode = node

    # Keep a record of the original nodes passed to this method to split marker nodes if needed.
    originalStartNode = startNode
    originalEndNode = endNode

    # Extract content based on block-level nodes (paragraphs and tables). Traverse through parent nodes to find them.
    # We will split the first and last nodes' content, depending if the marker nodes are inline.
    startNode = get_ancestor_in_body(startNode)
    endNode = get_ancestor_in_body(endNode)

    isExtracting = True
    isStartingNode = True
    # The current node we are extracting from the document.
    currNode = startNode

    # Begin extracting content. Process all block-level nodes and specifically split the first
    # and last nodes when needed, so paragraph formatting is retained.
    # Method is a little more complicated than a regular extractor as we need to factor
    # in extracting using inline nodes, fields, bookmarks, etc. to make it useful.
    while (isExtracting) :
        
        # Clone the current node and its children to obtain a copy.
        cloneNode = currNode.clone(True)
        isEndingNode = currNode == endNode

        if (isStartingNode or isEndingNode) :
            
            # We need to process each marker separately, so pass it off to a separate method instead.
            # End should be processed at first to keep node indexes.
            if (isEndingNode) :
                # !isStartingNode: don't add the node twice if the markers are the same node.
                process_marker(cloneNode, nodes, originalEndNode, currNode, isInclusive, False, not isStartingNode, False)
                isExtracting = False

            # Conditional needs to be separate as the block level start and end markers, maybe the same node.
            if (isStartingNode) :
                process_marker(cloneNode, nodes, originalStartNode, currNode, isInclusive, True, True, False)
                isStartingNode = False
            
        else :
            # Node is not a start or end marker, simply add the copy to the list.
            nodes.append(cloneNode)

        # Move to the next node and extract it. If the next node is None,
        # the rest of the content is found in a different section.
        if (currNode.next_sibling == None and isExtracting) :
            # Move to the next section.
            nextSection = currNode.get_ancestor(aw.NodeType.SECTION).next_sibling.as_section()
            currNode = nextSection.body.first_child
            
        else :
            # Move to the next node in the body.
            currNode = currNode.next_sibling
            
    # For compatibility with mode with inline bookmarks, add the next paragraph (empty).
    if (isInclusive and originalEndNode == endNode and not originalEndNode.is_composite) :
        include_next_paragraph(endNode, nodes)

    # Return the nodes between the node markers.
    return nodes