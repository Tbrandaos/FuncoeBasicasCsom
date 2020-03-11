using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.SharePoint.Client;
using System.Text;

namespace Teste
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        ClientContext context = new ClientContext("Site URL");

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                GetTitle();
            }
        }

        protected void GetTitle()
        {
            try
            {
                Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();

                lbl1.Text = web.Title;
            }
            catch (Exception ex)
            {
                lbl1.Text = ex.Message;
                throw;
            }
            
        }

        protected void GetTitleAndDescription()
        {
            try
            {
                Web web = context.Web;
                context.Load(web, w => w.Title, w => w.Description);
                context.ExecuteQuery();

                lbl1.Text = web.Title;
                lbl2.Text = web.Description;
            }
            catch (Exception ex)
            {
                lbl1.Text = ex.Message;
                throw;
            }
           
        }

        protected void WriteTitleAndDescription()
        {
            try
            {
                Web web = context.Web;
                web.Title = "Novo Teste";
                web.Description = "Descrição mudada via código";
                web.Update();
                context.ExecuteQuery();
                lbl1.Text = web.Title;
                lbl2.Text = web.Description;
            }
            catch (Exception ex)
            {
                lbl1.Text = ex.Message;
                throw;
            }
           
        }

        protected void NewSite()
        {
            try
            {
                WebCreationInformation creation = new WebCreationInformation();

                creation.Url = "web1";
                creation.Title = "Novo Site";
                Web newWeb = context.Web.Webs.Add(creation);

                context.Load(newWeb, w => w.Title);
                context.ExecuteQuery();

                lbl1.Text = newWeb.Title;
            }
            catch (Exception ex)
            {
                lbl1.Text = ex.Message;
                throw;
            }
            
        }

        protected void GetAllList()
        {
            try
            {
                Web web = context.Web;
                context.Load(web.Lists, lists => lists.Include(list => list.Title, list => list.Id));
                context.ExecuteQuery();

                foreach (List list in web.Lists)
                {
                    lbl1.Text = lbl1.Text + "-" + list.Title;
                    lbl2.Text = lbl2.Text + "-" + list.Id;
                }
            }
            catch (Exception ex)
            {
                lbl1.Text = ex.Message;
                throw;
            }
            
        }

        protected void CreateList()
        {
            try
            {
                Web web = context.Web;

                ListCreationInformation listCreation = new ListCreationInformation();
                listCreation.Title = "Lista Criada por Código";
                listCreation.TemplateType = (int)ListTemplateType.Announcements;
                List list = web.Lists.Add(listCreation);
                listCreation.Description = "Descrição da Lista";
                list.Update();
                context.ExecuteQuery();
                lbl1.Text = listCreation.Title;
                lbl2.Text = listCreation.Description;
            }
            catch (Exception ex)
            {
                lbl1.Text = ex.Message;
                throw;
            }
            
        }

        protected void DeleteList()
        {
            try
            {
                Web web = context.Web;

                List list = web.Lists.GetByTitle("Lista Criada por Código");
                list.DeleteObject();
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                lbl1.Text = ex.Message;
                throw;
            }
           
        }

        protected void GetListItems()
        {
            try
            {
                List testeList = context.Web.Lists.GetByTitle("Teste");
                CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
                Microsoft.SharePoint.Client.ListItemCollection items = testeList.GetItems(query);

                context.Load(items);
                context.ExecuteQuery();

                foreach (Microsoft.SharePoint.Client.ListItem li in items)
                {
                    lbl1.Text = lbl1.Text + ", " + li["Title"];
                }
            }
            catch (Exception ex)
            {
                lbl1.Text = ex.Message;
                throw;
            }
            
        }

        protected void CreateListItem()
        {
            try
            {
                List testeList = context.Web.Lists.GetByTitle("Teste");
                ListItemCreationInformation itemCreate = new ListItemCreationInformation();
                Microsoft.SharePoint.Client.ListItem newItem = testeList.AddItem(itemCreate);
                newItem["Title"] = "Paulo";
                newItem["Sobrenome"] = "Almeida";
                newItem.Update();

                context.ExecuteQuery();
                lbl1.Text = "Item Criado com sucesso";
            }
            catch (Exception ex)
            {
                lbl1.Text = ex.Message;
                throw;
            }           
        }

        protected void UpdateListItem()
        {
            try
            {
                List testeList = context.Web.Lists.GetByTitle("Teste");
                Microsoft.SharePoint.Client.ListItem item = testeList.GetItemById(1);
                item["Title"] = "Nome alterado via código";
                item.Update();
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                lbl1.Text = ex.Message;
                throw;
            }
            
        }

        protected void DeleteListItem()
        {
            try
            {
                List testeList = context.Web.Lists.GetByTitle("Teste");
                Microsoft.SharePoint.Client.ListItem item = testeList.GetItemById(6);
                item.DeleteObject();
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                lbl1.Text = ex.Message;
                throw;
            }
            
        }

    }
}