from pymongo import MongoClient
from datetime import datetime
from bson.objectid import ObjectId

# criação da conexão com o MongoDB
connection_string = "mongodb://pedro:1234@localhost:27017/?authSource=admin"
client = MongoClient(connection_string)
db_connection = client["bemp"]
collection = db_connection["pedidos"]

# funções para manipulação dos pedidos
def insert_order(args: dict):

    current_year = datetime.now().year
    last_order = collection.find_one(sort=[("pedido", -1)])

    if last_order and isinstance(last_order.get("pedido"), str) and "-" in last_order["pedido"]:
        try:
            last_order_number = int(last_order["pedido"].split("-")[1])
            new_order_number = last_order_number + 1
        except Exception:
            new_order_number = 1
    else:
        new_order_number = 1

    doc = {
        "pedido": f"{current_year}-{new_order_number:03d}",
        "cliente": args.get("cliente"),
        "data": args.get("data", datetime.now()),
        "endereco": args.get("endereco"),
        "telefone": args.get("telefone"),
        "tipo_pedido": args.get("tipo_pedido"),
        "categoria": args.get("categoria"),
        "subcategoria": args.get("subcategoria"),
        "especificacao": args.get("especificacao"),
        "valor": args.get("valor", {"tipo": "avista", "valor_total": 0}),
        "parceiros": args.get("parceiros", {"tipo": "sem_parceiros", "lista": []}),
        "status": args.get("status", "pendente")
    }

    # Converte a data caso esteja como string
    if isinstance(doc["data"], str):
        try:
            doc["data"] = datetime.strptime(doc["data"], "%d/%m/%Y")
        except:
            doc["data"] = datetime.now()

    # Converte datas dentro de parceiros se necessário
    parceiros = doc.get("parceiros", {})
    if isinstance(parceiros, dict) and "lista" in parceiros:
        for p in parceiros["lista"]:
            if "data_pagamento" in p and isinstance(p["data_pagamento"], str):
                try:
                    p["data_pagamento"] = datetime.strptime(p["data_pagamento"], "%d/%m/%Y")
                except:
                    p["data_pagamento"] = None

    inserted = collection.insert_one(doc)
    return inserted.inserted_id

#  função para atualizar um pedido existente
def update_order(pedido_identifier: str, updates: dict):
    """
    Atualiza um pedido identificado por seu campo 'pedido' (ex: '2025-002').
    'updates' deve conter os campos a alterar (mesmo formato do documento).
    Retorna dict com matched_count e modified_count.
    """
    if not pedido_identifier:
        raise ValueError("pedido_identifier é necessário")

    # Tantar conversão se data for fornecida como string
    if "data" in updates and isinstance(updates["data"], str):
        try:
            updates["data"] = datetime.strptime(updates["data"], "%d/%m/%Y")
        except:
            updates["data"] = datetime.now()

    # Converter datas nos parceiros se existirem
    if "parceiros" in updates and isinstance(updates["parceiros"], dict):
        lista = updates["parceiros"].get("lista", [])
        for p in lista:
            if "data_pagamento" in p and isinstance(p["data_pagamento"], str):
                try:
                    p["data_pagamento"] = datetime.strptime(p["data_pagamento"], "%d/%m/%Y")
                except:
                    p["data_pagamento"] = None

    res = collection.update_one({"pedido": pedido_identifier}, {"$set": updates})
    return {"matched": res.matched_count, "modified": res.modified_count}

#  função para consultar pedidos
def get_orders(filtro: dict):
    filtro = filtro or {}
    # valida se o filtro já é um filtro do Mongo
    try:
        if any(k.startswith('$') for k in filtro.keys()):
            query = filtro
        elif "q" in filtro:
            termo = filtro["q"].strip()
            # busca por número de pedido exato
            if termo.count('-') == 1 and termo.split('-')[0].isdigit():
                query = {"pedido": termo}
            else:
                regex_clause = {"$regex": termo, "$options": "i"}
                query = {
                    "$or": [
                        {"cliente": regex_clause},
                        {"pedido": regex_clause},
                        {"categoria": regex_clause},
                        {"subcategoria": regex_clause},
                        {"especificacao": regex_clause},
                        {"endereco": regex_clause},
                        {"tipo_pedido": regex_clause},
                        {"status": regex_clause},
                        {"parceiros.lista.nome": regex_clause},
                        {"parceiros.lista.telefone": regex_clause}
                    ]
                }
        else:
            # transformar filtros simples
            query = {}
            if "cliente" in filtro:
                query["cliente"] = {"$regex": filtro["cliente"], "$options": "i"}
            if "pedido" in filtro:
                query["pedido"] = filtro["pedido"]
            if "categoria" in filtro:
                query["categoria"] = filtro["categoria"]
            if "tipo_pedido" in filtro:
                query["tipo_pedido"] = filtro["tipo_pedido"]
            # filtros por data
            if "date_from" in filtro or "date_to" in filtro:
                date_query = {}
                if "date_from" in filtro:
                    try:
                        df = datetime.strptime(filtro["date_from"], "%d/%m/%Y")
                        date_query["$gte"] = df
                    except:
                        pass
                if "date_to" in filtro:
                    try:
                        dt = datetime.strptime(filtro["date_to"], "%d/%m/%Y")
                        date_query["$lte"] = dt
                    except:
                        pass
                if date_query:
                    query["data"] = date_query
    except Exception:
        query = {}

    resultados = list(collection.find(query))
    return resultados
