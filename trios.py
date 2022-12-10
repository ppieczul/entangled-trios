# Copyright 2022 oldcrap.org
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#    http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
import os
import sys
import pandas
from itertools import combinations


def remove_unused_columns(data):
    for i in list(data.keys()):
        if "YES" not in i:
            del data[i]


class Relation:
    def closer(self):
        self.proximity += 1

    def __init__(self, other):
        self.other = other
        self.proximity = 0


class Person:
    def similar_to(self, other):
        if other.name is not self.name:
            self.relations[other.name].closer()

    def sorted_relations(self):
        return sorted(self.relations.values(), key=lambda x: x.get_proximity())

    def proximity_factor(self):
        return sum(map(lambda x: x.proximity, self.relations.values()))

    def __init__(self, name, index, others):
        self.name = name
        self.index = index
        self.relations = {}
        for other in others:
            if other is not self.name:
                self.relations[other] = Relation(other)


class Population:
    def apply_similarities(self, cols):
        for col in cols.values():
            for i in self.population.values():
                val = col[i.index]
                for j in self.population.values():
                    if val == col[j.index] and val != '':
                        j.similar_to(i)

    def sort_by_proximity(self):
        return {key: v for key, v in
                sorted(self.population.items(), key=lambda x: x[1].proximity_factor())}

    def __init__(self, participants):
        self.population = {}
        for i, name in participants:
            self.population[name] = Person(name, i, names)


def build_trios(population):
    result = []
    left = []
    while len(population) > 0:
        p1 = population.popitem()[1]
        rest = list(combinations(population.values(), 2))
        options = []
        for r in rest:
            if len(r) == 2:
                p2 = r[0]
                p3 = r[1]
                f12 = p1.relations[p2.name].proximity
                f21 = p2.relations[p1.name].proximity
                assert (f12 == f21)
                f13 = p1.relations[p3.name].proximity
                f31 = p3.relations[p1.name].proximity
                assert (f13 == f31)
                f23 = p2.relations[p3.name].proximity
                f32 = p3.relations[p2.name].proximity
                assert (f23 == f32)
                factor = f12 + f13 + f23
                options.append((factor, p1.name, p2.name, p3.name))
                if factor == 0:
                    break
        if len(options) > 0:
            s = min(options, key=lambda x: x[0])
            result.append(s)
            population.pop(s[2])
            population.pop(s[3])
        else:
            left.append(p1)
    return result, left


if __name__ == '__main__':
    if len(sys.argv) <= 2:
        print('Usage:', os.path.basename(sys.argv[0]), '[input.xlsx] [output.xlsx]')
        exit(1)
    filename = sys.argv[1]
    writename = sys.argv[2]
    try:
        df_in = pandas.read_excel(filename, sheet_name='People')
        writer = pandas.ExcelWriter(writename)

        columns = df_in.to_dict()
        names_dict = columns['Full Name']
        names = list(names_dict.values())
        remove_unused_columns(columns)

        people = Population(names_dict.items())
        people.apply_similarities(columns)
        population_by_proximity = people.sort_by_proximity()
        initial_size = len(population_by_proximity)

        trios, leftover = build_trios(population_by_proximity)

        df_out = pandas.DataFrame(columns=['Person 1', 'Person 2', 'Person 3', 'Proximity'])
        df_in.set_index('Full Name', inplace=True)
        for idx, trio in enumerate(trios):
            print(trio)
            df_out.loc[idx + 1, 'Proximity'] = trio[0]
            df_out.loc[idx + 1, 'Person 1'] = trio[1]
            df_out.loc[idx + 1, 'Person 2'] = trio[2]
            df_out.loc[idx + 1, 'Person 3'] = trio[3]
            for p in range(1, 4):
                df_in.loc[trio[p], 'Trio'] = idx + 1
                df_in.loc[trio[p], 'Trio Proximity'] = trio[0]
                df_in.loc[trio[p], 'Personal Proximity'] = people.population[trio[p]].proximity_factor()

        print('Number of people: ', initial_size)
        print('Number of trios:  ', len(trios))
        print('People left:      ', len(leftover))

        for person in leftover:
            df_in.loc[person.name, 'Trio'] = ''
            print('Left: ', person.name)
            df_in.loc[person.name, 'Personal Proximity'] = people.population[person.name].proximity_factor()

        df_in.to_excel(writer, sheet_name='People with Trios')
        df_out.to_excel(writer, sheet_name='Trios')
        writer.close()

    except (FileNotFoundError, ValueError) as exception:
        print(exception)
        exit(1)
