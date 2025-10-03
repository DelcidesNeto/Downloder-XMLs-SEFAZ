import json
def criarJson():
    jsonInicial = {'links': []}
    with open('links.json','wt+') as arqivoJson:
        json.dump(jsonInicial, arqivoJson, indent=4)    


def adicionarLinks(link):
    with open('links.json', 'r', encoding='utf-8') as arquivoJson:
        linksJson = json.load(arquivoJson)
        linksJson['links'].append(link)
        with open('links.json','wt+') as links:
            json.dump(linksJson, links, indent=4)
criarJson()