from itertools import groupby

class Edge(object):
    def __init__(self, start, end, weight):
        self.startNodeId = start
        self.endNodeId = end
        self.weight = weight


class Node(object):
    def __init__(self, nodeId, edgeList):
        self.nodeId = nodeId
        self.edgeList = edgeList


class Path(object):
    def __init__(self, curNodeId):
        self.visited = False
        self.weight = 100000000
        self.curNodeId = curNodeId
        self.routeList = []

class Graph(object):
    def __init__(self, nodeList):
        self.nodeList = nodeList
        self.pathDic = {}

    def initPaths(self, originNodeId):
        self.pathDic = {}
        originNode = None
        for node in self.nodeList:
            if node.nodeId == originNodeId:
                originNode = node
            self.pathDic[node.nodeId] = Path(node.nodeId)

        if originNode is None:
            print("originNode is none")
            return
        else:
            for edge in originNode.edgeList:
                path = self.pathDic[edge.endNodeId]
                if path is None:
                    print("path is None")
                    return
                else:
                    path.weight = edge.weight
                    path.routeList.append(originNodeId)

        path = self.pathDic[originNodeId]

        path.weight = 0
        path.visit = True
        path.routeList.append(originNodeId)

    def getMinPath(self, originNodeId):
        destNode = None
        weight = 100000000
        for node in self.nodeList:
            path = self.pathDic[node.nodeId]
            if path.visited == False and path.weight < weight:
                weight = path.weight
                destNode = node

        return destNode

    def dijkstra(self, originNodeId, destNodeId):

        if not destNodeId:
            return
        else:
            self.initPaths(originNodeId)
            curNode = self.getMinPath(originNodeId)
            while curNode is not None:
                curPath = self.pathDic[curNode.nodeId]
                curPath.visited = True

                for edge in curNode.edgeList:
                    minPath = self.pathDic[edge.endNodeId]

                    if curNode.nodeId not in curPath.routeList:

                        if minPath.weight >(curPath.weight + edge.weight):
    
                            minPath.weight = curPath.weight + edge.weight
                            minPath.routeList = curPath.routeList + [curNode.nodeId]
    
                        elif minPath.weight == (curPath.weight + edge.weight):
    
                            minPath.weight = curPath.weight + edge.weight
                            minPath.routeList += curPath.routeList + [curNode.nodeId]

                curNode = self.getMinPath(originNodeId)

            route = self.pathDic[destNodeId].routeList

            weight = self.pathDic[destNodeId].weight

            if route == []:
                return
            else:

                if len(route) == 1:
                    route = [route[i:i + (len(route) // route.count(originNodeId))] for i in
                              range(0, len(route), len(route) // route.count(originNodeId))]
                    
                elif [[originNodeId] + (list(g)) for k, g in groupby(route, lambda x: x == originNodeId) if not k] == []:
                    route = [route[i:i + (len(route) // route.count(originNodeId))] for i in
                             range(0, len(route), len(route) // route.count(originNodeId))]
                else:

                    route = [[originNodeId]+(list(g)) for k, g in groupby(route, lambda x: x == originNodeId) if not k]

                return originNodeId,destNodeId,route,weight




