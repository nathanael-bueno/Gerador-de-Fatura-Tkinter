from tkinter import *
from tkinter import ttk, messagebox
from docxtpl import DocxTemplate
import os
import datetime

def add_item():
    try:
        quant = int(ent_quantidade.get().strip())
        preco = float(f"{float(ent_preco.get().strip()):.2f}")
        if ent_produto.get().strip() == "" or quant <= 0 or preco <= 0:
            messagebox.showinfo(title="Erro", message="Você não pode adicionar um novo item com Valores Invalidos, verifique se os dados de: Produto, Quantidade e Preço")
            return
        produto = (ent_quantidade.get(), ent_produto.get(), ent_preco.get(), str(quant * preco))
        fatura_list.append(produto)
        tv_produtos.insert("", "end", values=produto)
        ent_produto.delete(0, END)
        ent_quantidade.delete(0, END)
        ent_preco.delete(0, END)
        ent_quantidade.insert(0, "1")
        ent_preco.insert(0, "0.0")
        ent_produto.focus()
    except Exception as error:
        print(error)
        messagebox.showerror(title="Erro", message="Ocorreu um Erro ao tentar Adicionar um Novo Item!")

def remover_item():
    try:
        item_selecionado = tv_produtos.selection()[0]
        fatura_list.remove(tv_produtos.item(item_selecionado, "values"))
        tv_produtos.delete(item_selecionado)
    except Exception as error:
        print(error)
        messagebox.showerror(title="Erro", message="Ocorreu um Erro ao tentar Remover um Item!")

def zerar_fatura():
    try:
        ent_nome.delete(0, END)
        ent_telefone.delete(0, END)
        ent_produto.delete(0, END)
        ent_quantidade.delete(0, END)
        ent_preco.delete(0, END)
        ent_desconto.delete(0, END)
        ent_desconto.insert(0, "0")
        ent_quantidade.insert(0, "1")
        ent_preco.insert(0, "0.0")
        ent_nome.focus()
        tv_produtos.delete(*tv_produtos.get_children())
        fatura_list.clear()
    except Exception as error:
        print(error)
        messagebox.showerror(title="Erro", message="Ocorreu um Erro ao tentar Zerar a Fatura!")

def salvar_fatura():
    try:
        if ent_nome.get().strip() == "" or ent_telefone.get().strip() == "":
            messagebox.showinfo(title="Erro", message="Você não pode Gerar uma Nova Fatura sem preecher todos os Dados Necessários: Nome e Telefone!")
            return
        total = 0
        for c in fatura_list:
            total += float(c[3])
        pasta = os.path.dirname(__file__)
        doc = DocxTemplate(pasta + "/fatura_template.docx")
        doc.render({"nome":ent_nome.get(),
                    "tel":ent_telefone.get(),
                    "fatura_list":fatura_list,
                    "subtotal":str(total),
                    "desconto":ent_desconto.get(),
                    "total":str(total-(total*float(ent_desconto.get())/100))})
        nome_doc = pasta + "/" + "fatura-" + ent_nome.get().strip().replace(" ", "-") + "-" + datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S") + ".docx"
        doc.save(nome_doc)
        zerar_fatura()
        messagebox.showinfo(title="Salvo", message="Sua Fatura foi Salva com Sucesso")
    except Exception as error:
        print(error)
        messagebox.showerror(title="Erro", message="Ocorreu um Erro ao tentar Gerar o Documento da Fatura")



app = Tk()
app.title("App do Bueno")
app.geometry("1000x450")
app.config()

tela = Frame(app)
tela.pack(padx=10, pady=10, fill="both", expand=True)


lb_nome = Label(tela, text="Nome")
lb_nome.grid(row=0, column=0)
ent_nome = Entry(tela)
ent_nome.grid(row=1, column=0, pady=2)

lb_telefone = Label(tela, text="Telefone")
lb_telefone.grid(row=2, column=0, pady=2)
ent_telefone = Entry(tela)
ent_telefone.grid(row=3, column=0)

lb_produto = Label(tela, text="Produto")
lb_produto.grid(row=0, column=1)
ent_produto = Entry(tela)
ent_produto.grid(row=1, column=1, pady=2)

lb_quantidade = Label(tela, text="Quantidade")
lb_quantidade.grid(row=2, column=1, pady=2)
ent_quantidade = Spinbox(tela, from_=1, to=500, increment=1)
ent_quantidade.grid(row=3, column=1)

lb_preco = Label(tela, text="Preço da Unidade")
lb_preco.grid(row=0, column=2)
ent_preco = Spinbox(tela, from_=0.0, to=10000, increment=0.5)
ent_preco.grid(row=1, column=2)

lb_desconto = Label(tela, text="Desconto Total(%)")
lb_desconto.grid(row=2, column=2)
ent_desconto = Spinbox(tela, from_=0, to=100, increment=1)
ent_desconto.grid(row=3, column=2)


colunas = ("quant", "nome", "preco", "total")
tv_produtos = ttk.Treeview(tela, columns=colunas, show="headings")
tv_produtos.grid(row=5, column=0, padx=5, pady=50, columnspan=3, rowspan=2)
tv_produtos.heading('quant', text="Quantidade")
tv_produtos.heading('nome', text="Nome")
tv_produtos.heading('preco', text="Preco por Unidade")
tv_produtos.heading('total', text="Total")
fatura_list = []


btn_adicionar = Button(tela, text="Adicionar Item", command=add_item)
btn_adicionar.grid(row=0, column=3, padx=5, pady=5, rowspan=2)
btn_remover = Button(tela, text="Remover Item", command=remover_item)
btn_remover.grid(row=2, column=3, padx=5, pady=5, rowspan=2)
btn_fatura = Button(tela, text="Criar Fatura", command=salvar_fatura)
btn_fatura.grid(row=4, column=3, padx=5, pady=5, rowspan=2)
btn_zerar = Button(tela, text="Zerar Fatura", command=zerar_fatura)
btn_zerar.grid(row=5, column=3, padx=5, pady=5, rowspan=2)

# Resolver problema de arquivo corrompido

app.mainloop()