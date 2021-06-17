from ortools.graph import pywrapgraph

ACCURACY = 3

def construct_flow(student_prefs, prof_prefs, course_info):

    N, M = len(student_prefs.keys()), sum([target for target, _ in course_info.values()])
    source, sink = N + M, N + M + 1
    flow = pywrapgraph.SimpleMinCostFlow(2 + N + M)

    # Each student cannot fill >1 course slot
    for i in range(N):
        flow.AddArcWithCapacityAndUnitCost(source, i, 1, 0)

    # Each course slot cannot have >1 TA
    # Value of filling slot is reciprocal with slot index
    node = N
    for target, value in course_info.values():
        total = sum(1/i for i in range(1, target + 1))
        for i in range(1, target + 1):
            flow.AddArcWithCapacityAndUnitCost(node, sink, 1, int(-value / i / total * 10 ** ACCURACY))
            node += 1

    # Edge weights given by preferences
    for i, student in enumerate(student_prefs.keys()):
        node = N
        for course, target, _ in course_info.items():
            if course in student_prefs[student] and student in prof_prefs[course]:
                cost = -(student_prefs[student][course] + prof_prefs[course][student]) * 10 ** ACCURACY
                for slot in range(target):
                    flow.AddArcWithCapacityAndUnitCost(i, node + slot, 1, cost)
            node += target       

    # Every student must TA a course
    flow.SetNodeSupply(source, N)
    flow.SetNodeSupply(sink, -N)

    return flow

# if __name__ == "__main__":

#     M = 25
#     capacities = np.random.randint(1, 5, M)
#     N = np.sum(capacities)
#     prefs1 = np.tile(np.arange(M), (N, 1))
#     prefs2 = np.tile(np.arange(N), (M, 1))
#     for i in range(N):
#         prefs1[i, 0:np.random.randint(0, 0.67 * M)] = 0
#         np.random.shuffle(prefs1[i])
#     for k in range(M):
#         prefs2[k, 0:np.random.randint(0, 0.67 * N)] = 0
#         np.random.shuffle(prefs2[k])

#     flow = construct_flow(prefs1, prefs2, capacities)
    
#     if flow.Solve() == flow.OPTIMAL:
#         print('Max value: ', -flow.OptimalCost())
#         print('')
#         print('Matches: ')
#         print('  Arc    Student pref   Prof pref')
#         for i in range(flow.NumArcs()):
#             cost = flow.Flow(i) * flow.UnitCost(i)
#             t, h = flow.Tail(i), flow.Head(i) - N 
#             if cost < 0:
#                 print(f'{t:2} -> {h:2}     {prefs1[t, h]:2}             {prefs2[h, t]:2}')