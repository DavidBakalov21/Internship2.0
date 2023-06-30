from collections import deque


class Node:
    def __init__(self, left, right, value):
        self.left=left
        self.right=right
        self.value=value

L3node=Node(left=None, right=None, value=1000)
R2node=Node(left=None, right=None, value=76)
L2node=Node(left=L3node, right=R2node, value=67)
Lnode=Node(left=L2node, right=None, value=9)
Rrnode=Node(left=None, right=None, value=112)
RNode=Node(left=None, right=Rrnode, value=6)
root=Node(left=Lnode, right=RNode, value=1)

def traverseInorder(Node):
    if Node is None:
        return
    traverseInorder(Node.left)
    print(Node.value)
    traverseInorder(Node.right)


def inorder_traversal(root):
    res = []
    if not root:
        return res
    stack = []
    curr_node = root
    while curr_node or stack:
        while curr_node:
            stack.append(curr_node)
            curr_node = curr_node.left
        curr_node = stack.pop()
        res.append(curr_node.value)
        curr_node = curr_node.right

    return res
def BFS(root):
    if root is None:
        return

    queue = deque()
    queue.append(root)
    while queue:
        node = queue.popleft()
        print(node.value)

        if node.left is not None:
            queue.append(node.left)
        if node.right is not None:
            queue.append(node.right)

def DFS(root):
    if root is None:
        return
    print(root.value)  # Or perform any other operation with the node

    DFS(root.left)
    DFS(root.right)
traverseInorder(root)
print("------")
print(inorder_traversal(root))
print("------")
BFS(root)
print("---")
DFS(root)