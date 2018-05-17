using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint;
using System.Net.Mail;
using System.Net;
using jobContratosFimVigencia.PSS.ListData;
using System.Data.Services.Client;
using System.Windows.Forms;

namespace jobContratosFimVigencia
{
    public class MonitoraJob : SPJobDefinition{
        //define listas e variáveis
        List<mdContrato> ListaResultContrato = null;
        List<mdHistoricoEnvioEmail> ListaResultHistorico = null;
        mdConfig itemConfig = new mdConfig();
        string pMsgErroMail = string.Empty;
        string vURI = string.Empty;


        public MonitoraJob() : base() { }

        public MonitoraJob(string jobName, SPService service, SPServer server, SPJobLockType targetType) : 
            base(jobName, service, server, targetType) { }

        public MonitoraJob(string jobName, SPWebApplication webApplication) : base(jobName, webApplication, null, SPJobLockType.ContentDatabase) 
            { this.Title = "B2W Job - Contratos fim de vigência"; }

        public override void Execute(Guid contentDbId)
        {
            vURI = itemConfig.siteContratos + itemConfig.uri;
            
            //verifica os contratos com a data de fim de vigência para os próximos 60 dias e 
            //envia um email aos responsáveis (lista de NotificaçãoEmail, responsável e contato B2W)
            VerificaContratosFimVigencia();

            //verifica a data de Notificação de rescisão dos contratos e 
            //envia um email aos responsáveis (lista de NotificaçãoEmail, responsável e contato B2W)
            VerificaContratosAvisoCancelamento();
        }


        #region ===== FUNÇÕES =====
        public void VerificaContratosFimVigencia()
        {
            ListaResultHistorico = new List<mdHistoricoEnvioEmail>();

            ContratosDataContext dc = new ContratosDataContext(new Uri(vURI));
            dc.Credentials = new NetworkCredential(itemConfig.user, itemConfig.password, itemConfig.domain);

            //Define mensagem do email
            string vSubject = "Gestão de Contratos - Aviso de Fim de Vigência";
            string vBodyInicio = "<span style='font-family:Arial; font-size:14px'>Contratos próximos do fim de vigência.</span><br/><br/>" +
                                "<table width=100% style='font-family:Arial; font-size:14px'><tr style='background-color:#EEE9E9; height:25px'>" +
                                "<td width=40px>N.Documento</td><td width=40px>Fornecedor</td>" +
                                "<td width=40px>Responsável</td><td width=40px>Data fim da vigência</td>" +
                                "<td width=*>&nbsp;</td></tr>";
            string vBodyMeio = "<tr style='height:25px'><td width=40px>{0}</td><td width=40px>{1}</td>" +
                                "<td width=40px>{2}</td><td width=40px>{3}</td><td width=*><a href='{4}'>Ver contrato</a></td></tr>";
            string vBodyFim = "</table>";
            string vBody = string.Empty;
          
            #region ============ Gera lista de Contratos com fim de vigência para os próximos 60 dias, caso o alerta de email esteja ativado ============
            
            ListaResultContrato = new List<mdContrato>();
            mdContrato itemContrato = null;

            var queryContratos = (from lstContratos in dc.Contratos
                                  where lstContratos.SituaçãoValue.Equals("Ativo") &&
                                  lstContratos.AlertaDeEmail == true &&      
                                  //lstContratos.VigênciaFim >= DateTime.Today &&
                                  //lstContratos.VigênciaFim <= DateTime.Today.Add(TimeSpan.FromDays(60))
                                  lstContratos.DataNotifVigencia.Equals(DateTime.Today)
                                  select new
                                  {
                                      NumDocto = lstContratos.NúmeroDoDocumento,
                                      Fornecedor = lstContratos.Fornecedor,
                                      Responsavel = lstContratos.Responsável,
                                      ContatoB2W = lstContratos.ContatoB2W,
                                      VigenciaFim = lstContratos.VigênciaFim,
                                      ID = lstContratos.ID
                                  });

            foreach (var itemQuery in queryContratos)
            {
                itemContrato = new mdContrato();

                itemContrato.NumDocto = itemQuery.NumDocto.ToString();
                itemContrato.Fornecedor = itemQuery.Fornecedor.NomeFantasia.ToString();
                itemContrato.Responsavel = itemQuery.Responsavel.WorkEMail.ToString();
                itemContrato.ContatoB2W = itemQuery.ContatoB2W.WorkEMail.ToString();
                itemContrato.VigenciaFim = itemQuery.VigenciaFim.Value.ToString("d/MM/yyy");
                itemContrato.LinkContrato = itemConfig.siteContratos + itemConfig.linkContratoView + "?ID=" + itemQuery.ID;
                
                ListaResultContrato.Add(itemContrato);
            }
            #endregion

            
            #region ============ Envia email para Responsável e Contato B2W de cada contrato ============
            
            if (ListaResultContrato.Count > 0)
            {
                bool emailSent;

                //Monta mensagem com contratos
                foreach (var itemContratos in ListaResultContrato)
                {
                    pMsgErroMail = string.Empty;

                    vBody = vBodyInicio +
                            String.Format(vBodyMeio, itemContratos.NumDocto, itemContratos.Fornecedor, itemContratos.Responsavel, itemContratos.VigenciaFim, itemContratos.LinkContrato) +
                            vBodyFim;

                    emailSent = SendMail(vSubject, vBody, itemContratos.Responsavel, itemContratos.ContatoB2W, null);

                    mdHistoricoEnvioEmail itemHistorico = new mdHistoricoEnvioEmail();

                    itemHistorico.Origem = "GestaoContratos";
                    itemHistorico.Para = itemContratos.Responsavel + "; " + itemContratos.ContatoB2W;
                    itemHistorico.Data = DateTime.Now.ToString("d/MM/yyy HH:mm");

                    if (emailSent)
                    {
                        itemHistorico.Status = "Sucesso";
                        itemHistorico.Mensagem = "Aviso de Fim de Vigência - E-mail enviado com sucesso!";
                    }
                    else
                    {
                        itemHistorico.Status = "Erro";
                        itemHistorico.Mensagem = "Aviso de Fim de Vigência - E-mail não enviado devido a falha: " + pMsgErroMail;
                    }

                    ListaResultHistorico.Add(itemHistorico);
                }
            }
            else
            {
                mdHistoricoEnvioEmail itemHistorico = new mdHistoricoEnvioEmail();

                itemHistorico.Origem = "GestaoContratos";
                itemHistorico.Para = string.Empty;
                itemHistorico.Data = DateTime.Now.ToString("d/MM/yyy HH:mm");

                itemHistorico.Status = "Sucesso";
                itemHistorico.Mensagem = "Aviso de Fim de Vigência - Não possui contratos nas condições ou usuários cadastrados. (" + ListaResultContrato.Count + " contrato(s))";

                ListaResultHistorico.Add(itemHistorico);
            }
            
            #endregion
            

            #region ============ Envia email para Grupo de email ============

            vBody = string.Empty;

            if (ListaResultContrato.Count > 0)
            {
                bool emailSent;

                //Monta mensagem com contratos
                vBody = vBodyInicio;

                foreach (var itemContratos in ListaResultContrato)
                {
                    vBody += String.Format(vBodyMeio, itemContratos.NumDocto, itemContratos.Fornecedor, itemContratos.Responsavel, itemContratos.VigenciaFim, itemContratos.LinkContrato);
                }
                
                vBody += vBodyFim;

                //Monta lista com emails 
                var queryEmail = from email in dc.GrupoEmail
                                 where email.Acesso.Equals("Contratos - Fim de Vigência")
                                 select new
                                 {
                                     pEmail = email.Usuários
                                 };
                
                foreach (var itemMail in queryEmail)
                {
                    for (int x = 0; x < itemMail.pEmail.Count; x++)
                    {
                        pMsgErroMail = string.Empty;

                        emailSent = SendMail(vSubject, vBody, itemMail.pEmail[x].WorkEMail.ToString(), null, null);

                        mdHistoricoEnvioEmail itemHistorico = new mdHistoricoEnvioEmail();

                        itemHistorico.Origem = "GestaoContratos";
                        itemHistorico.Para = itemMail.pEmail[x].WorkEMail.ToString();
                        itemHistorico.Data = DateTime.Now.ToString("d/MM/yyy HH:mm");

                        if (emailSent)
                        {
                            itemHistorico.Status = "Sucesso";
                            itemHistorico.Mensagem = "Aviso de Fim de Vigência- E-mail enviado com sucesso!";
                        }
                        else
                        {
                            itemHistorico.Status = "Erro";
                            itemHistorico.Mensagem = "Aviso de Fim de Vigência - E-mail não enviado devido a falha: " + pMsgErroMail;
                        }

                        ListaResultHistorico.Add(itemHistorico);
                    }
                }
            }
            else
            {
                mdHistoricoEnvioEmail itemHistorico = new mdHistoricoEnvioEmail();

                itemHistorico.Origem = "GestaoContratos";
                itemHistorico.Para = string.Empty;
                itemHistorico.Data = DateTime.Now.ToString("d/MM/yyy HH:mm");

                itemHistorico.Status = "Sucesso";
                itemHistorico.Mensagem = "Aviso de Fim de Vigência- Não possui contratos nas condições ou usuários cadastrados. (" + ListaResultContrato.Count + " contrato(s))";

                ListaResultHistorico.Add(itemHistorico);            
            }

            #endregion


            GravaHistoricoEmailEnviado();
        }


        public void VerificaContratosAvisoCancelamento()
        {
            ListaResultHistorico = new List<mdHistoricoEnvioEmail>();

            ContratosDataContext dc = new ContratosDataContext(new Uri(vURI));
            dc.Credentials = new NetworkCredential(itemConfig.user, itemConfig.password, itemConfig.domain);

            //Define mensagem do email
            string vSubject = "Gestão de Contratos - Notificação de Rescisão";
            string vBodyInicio = "<span style='font-family:Arial; font-size:14px'>Contratos próximos do fim de vigência para notificação de rescisão.</span><br/><br/>" +
                                "<table width=100% style='font-family:Arial; font-size:14px'><tr style='background-color:#EEE9E9; height:25px'>" +
                                "<td width=40px>N.Documento</td><td width=40px>Fornecedor</td>" +
                                "<td width=40px>Responsável</td><td width=40px>Data fim da vigência</td>" +
                                "<td width=40px>Notif. Rescisão</td><td width=*>&nbsp;</td></tr>";
            string vBodyMeio = "<tr style='height:25px'><td width=40px>{0}</td><td width=40px>{1}</td>" +
                                "<td width=40px>{2}</td><td width=40px>{3}</td><td width=40px>{4}</td><td width=*><a href='{5}'>Ver contrato</a></td></tr>";
            string vBodyFim = "</table>";
            string vBody = string.Empty;

            #region ============ Gera lista de Contratos com fim de vigência com base na notificação de rescisão, caso o alerta de email esteja ativado ============

            ListaResultContrato = new List<mdContrato>();
            mdContrato itemContrato = null;

            var queryContratos = (from lstNotifContratos in dc.Contratos
                                  where lstNotifContratos.SituaçãoValue.Equals("Ativo") &&
                                  lstNotifContratos.AlertaDeEmail.Equals(true) &&
                                  //lstNotifContratos.VigênciaFim.Equals(DateTime.Today.Add(TimeSpan.FromDays(120)))
                                  lstNotifContratos.DataNotifRescisao.Equals(DateTime.Today)
                                  select new
                                  {
                                      NumDocto = lstNotifContratos.NúmeroDoDocumento,
                                      Fornecedor = lstNotifContratos.Fornecedor,
                                      Responsavel = lstNotifContratos.Responsável,
                                      ContatoB2W = lstNotifContratos.ContatoB2W,
                                      VigenciaFim = lstNotifContratos.VigênciaFim,
                                      NotifRescisao = lstNotifContratos.NotificaçãoDeRescisão,
                                      ID = lstNotifContratos.ID
                                  });

            foreach (var itemQuery in queryContratos)
            {
                itemContrato = new mdContrato();

                itemContrato.NumDocto = itemQuery.NumDocto.ToString();
                itemContrato.Fornecedor = itemQuery.Fornecedor.NomeFantasia.ToString();
                itemContrato.Responsavel = itemQuery.Responsavel.WorkEMail.ToString();
                itemContrato.ContatoB2W = itemQuery.ContatoB2W.WorkEMail.ToString();
                itemContrato.VigenciaFim = itemQuery.VigenciaFim.Value.ToString("d/MM/yyy");
                itemContrato.NotifRescisao = itemQuery.NotifRescisao.ToString() + " dias";
                itemContrato.LinkContrato = itemConfig.siteContratos + itemConfig.linkContratoView + "?ID=" + itemQuery.ID;

                ListaResultContrato.Add(itemContrato);
            }
            #endregion


            #region ============ Envia email para Responsável e Contato B2W de cada contrato ============

            if (ListaResultContrato.Count > 0)
            {
                bool emailSent;

                //Monta mensagem com contratos
                foreach (var itemContratos in ListaResultContrato)
                {
                    pMsgErroMail = string.Empty;

                    vBody = vBodyInicio +
                            String.Format(vBodyMeio, itemContratos.NumDocto, itemContratos.Fornecedor, itemContratos.Responsavel, itemContratos.VigenciaFim, itemContratos.NotifRescisao, itemContratos.LinkContrato) +
                            vBodyFim;

                    emailSent = SendMail(vSubject, vBody, itemContratos.Responsavel, itemContratos.ContatoB2W, null);

                    mdHistoricoEnvioEmail itemHistorico = new mdHistoricoEnvioEmail();

                    itemHistorico.Origem = "GestaoContratos";
                    itemHistorico.Para = itemContratos.Responsavel + "; " + itemContratos.ContatoB2W;
                    itemHistorico.Data = DateTime.Now.ToString("d/MM/yyy HH:mm");

                    if (emailSent)
                    {
                        itemHistorico.Status = "Sucesso";
                        itemHistorico.Mensagem = "Notificação de Rescisão - E-mail enviado com sucesso!";
                    }
                    else
                    {
                        itemHistorico.Status = "Erro";
                        itemHistorico.Mensagem = "Notificação de Rescisão - E-mail não enviado devido a falha: " + pMsgErroMail;
                    }

                    ListaResultHistorico.Add(itemHistorico);
                }
            }
            else
            {
                mdHistoricoEnvioEmail itemHistorico = new mdHistoricoEnvioEmail();

                itemHistorico.Origem = "GestaoContratos";
                itemHistorico.Para = string.Empty;
                itemHistorico.Data = DateTime.Now.ToString("d/MM/yyy HH:mm");

                itemHistorico.Status = "Sucesso";
                itemHistorico.Mensagem = "Notificação de Rescisão - Não possui contratos nas condições ou usuários cadastrados. (" + ListaResultContrato.Count + " contrato(s))";

                ListaResultHistorico.Add(itemHistorico);
            }

            #endregion


            #region ============ Envia email para Grupo de email ============

            vBody = string.Empty;

            if (ListaResultContrato.Count > 0)
            {
                bool emailSent;

                //Monta mensagem com contratos
                vBody = vBodyInicio;

                foreach (var itemContratos in ListaResultContrato)
                {
                    vBody += String.Format(vBodyMeio, itemContratos.NumDocto, itemContratos.Fornecedor, itemContratos.Responsavel, itemContratos.VigenciaFim, itemContratos.NotifRescisao, itemContratos.LinkContrato);
                }

                vBody += vBodyFim;

                //Monta lista com emails 
                var queryEmail = from email in dc.GrupoEmail
                                 where email.Acesso.Equals("Contratos - Fim de Vigência")
                                 select new
                                 {
                                     pEmail = email.Usuários
                                 };

                foreach (var itemMail in queryEmail)
                {
                    for (int x = 0; x < itemMail.pEmail.Count; x++)
                    {
                        pMsgErroMail = string.Empty;

                        emailSent = SendMail(vSubject, vBody, itemMail.pEmail[x].WorkEMail.ToString(), null, null);

                        mdHistoricoEnvioEmail itemHistorico = new mdHistoricoEnvioEmail();

                        itemHistorico.Origem = "GestaoContratos";
                        itemHistorico.Para = itemMail.pEmail[x].WorkEMail.ToString();
                        itemHistorico.Data = DateTime.Now.ToString("d/MM/yyy HH:mm");

                        if (emailSent)
                        {
                            itemHistorico.Status = "Sucesso";
                            itemHistorico.Mensagem = "Notificação de Rescisão - E-mail enviado com sucesso!";
                        }
                        else
                        {
                            itemHistorico.Status = "Erro";
                            itemHistorico.Mensagem = "Notificação de Rescisão - E-mail não enviado devido a falha: " + pMsgErroMail;
                        }

                        ListaResultHistorico.Add(itemHistorico);
                    }
                }
            }
            else
            {
                mdHistoricoEnvioEmail itemHistorico = new mdHistoricoEnvioEmail();

                itemHistorico.Origem = "GestaoContratos";
                itemHistorico.Para = string.Empty;
                itemHistorico.Data = DateTime.Now.ToString("d/MM/yyy HH:mm");

                itemHistorico.Status = "Sucesso";
                itemHistorico.Mensagem = "Notificação de Rescisão - Não possui contratos nas condições ou usuários cadastrados. (" + ListaResultContrato.Count + " contrato(s))";

                ListaResultHistorico.Add(itemHistorico);
            }

            #endregion


            GravaHistoricoEmailEnviado();
        }



        public void GravaHistoricoEmailEnviado()
        {
            //grava na lista de historico de notificações os emails enviados.
            mdConfig itemConfig = new mdConfig();

            ContratosDataContext dc = new ContratosDataContext(new Uri(vURI));
            dc.Credentials = new NetworkCredential(itemConfig.user, itemConfig.password, itemConfig.domain);

            //cria os novos registros gerados
            HistóricoJobEmailEnviadoItem novoitem = null;

            foreach (mdHistoricoEnvioEmail item in ListaResultHistorico)
            {
                novoitem = new HistóricoJobEmailEnviadoItem();

                novoitem.Origem = item.Origem;
                novoitem.Status = item.Status;
                novoitem.Mensagem = item.Mensagem;
                novoitem.Para = item.Para;
                novoitem.DataDeEnvio = item.Data;

                dc.AddToHistóricoJobEmailEnviado(novoitem);
                dc.SaveChanges();
            }
        }

        public bool SendMail(string vSubject, string vBody, string vTo, string vCC, string vBCC)
        {
            bool mailSent = false;
            SmtpClient smtpClient = null;

            try
            {
                mdConfig itemConfig = new mdConfig();

                smtpClient = new SmtpClient();
                smtpClient.Host = itemConfig.mailhost;
                smtpClient.Credentials = new NetworkCredential(itemConfig.user, itemConfig.password, itemConfig.domain);

                string vFrom = itemConfig.mailfrom;

                MailMessage mailMessage = new MailMessage(vFrom, vTo, vSubject, vBody);
                if (!String.IsNullOrEmpty(vCC)) //cópia
                {
                    MailAddress CCAddress = new MailAddress(vCC);
                    mailMessage.CC.Add(CCAddress);
                }
                if (!String.IsNullOrEmpty(vBCC)) //cópia oculta
                {
                    MailAddress BCCAddress = new MailAddress(vBCC);
                    mailMessage.Bcc.Add(BCCAddress);
                }
                mailMessage.IsBodyHtml = true;

                //Envia email 
                smtpClient.Send(mailMessage);
                mailSent = true;
            }
            catch (Exception e)
            {
                pMsgErroMail = e.Message + " - " + e.InnerException;
                mailSent = false;
            }

            return mailSent;
        }

        #endregion
    }
}


