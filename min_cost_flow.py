from ortools.graph import pywrapgraph
from copy import copy

DIGITS = 3

def construct_flow(weights, course_info):

    N, M = weights.shape[0], sum([cap - filled for cap, filled, _ in course_info.values()])
    source, sink = N + M, N + M + 1
    flow = pywrapgraph.SimpleMinCostFlow(2 + N + M)

    # Each student cannot fill >1 course slot
    for i in range(N):
        flow.AddArcWithCapacityAndUnitCost(source, i, 1, 0)

    # Each course slot cannot have >1 TA
    # Value of filling slot is reciprocal with slot index
    node = N
    for cap, filled, value in course_info.values():
        for i in range(filled, cap):
            flow.AddArcWithCapacityAndUnitCost(node, sink, 1, -fill_value(cap, i, value))
            node += 1

    # Edge weights given by preferences
    for si in range(N):
        node = N
        for ci, info in enumerate(course_info.values()):
            target, filled, _ = info
            if weights[si, ci] > -5:
                cost = int(-weights[si, ci] * 10 ** DIGITS)
                for slot in range(target - filled):
                    flow.AddArcWithCapacityAndUnitCost(si, node + slot, 1, cost)
            node += target

    # Max number of slots must be filled
    flow.SetNodeSupply(source, min(N, M))
    flow.SetNodeSupply(sink, -min(N, M))

    return flow

def fill_value(cap, index, weight):
    return int(weight / (index + 1) * 10 ** DIGITS)