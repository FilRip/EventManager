using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

#nullable enable

namespace FilRip.Utils.Extensions
{
    /// <summary>
    /// Classe d'extension pour gérer les handlers des méthodes d'évènement
    /// </summary>
    public static class ExtensionsEvents
    {
        private const BindingFlags binding = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.IgnoreCase;

        private const string NomMagasinEventWpf = "EventHandlersStore";
        private const string NomProprieteEventWinForm = "Events";
        private const string NomMethodeRetourneHandlerMagasinEventWpf = "GetRoutedEventHandlers";
        private const string NomChampsEventForm = "EVENT_";
        private const string NomChampsEventControlWinForm = "Event";
        private const string NomChampsEventControlWpf = "Event";
        private readonly static string[] listeSuffixeNomChampsEventWinForm = new string[] { "CHANGED", "FIRED" };

        #region Copier events

        /// <summary>
        /// Copie les méthodes abonné aux events (handlers) d'un objet à l'autre.
        /// Ainsi lorsqu'on recréer l'objet dest on récupère les handlers, pas besoin de se réabonner
        /// </summary>
        /// <param name="source">Objet source contenant les events abonnés</param>
        /// <param name="dest">Objet destination ou l'on va mettre les events abonné</param>
        /// <param name="eventName">Copie seulement les handlers d'un event en particulier (ou null/vide, pour tous les events)</param>
        /// <remarks>Les deux objets (source et dest) doivent être obligatoirement du même type</remarks>
        public static void CopyEventsTo(object source, object dest, string? eventName = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (dest == null) throw new ArgumentNullException(nameof(dest));

            Type typeSource = source.GetType();
            Type typeDest = dest.GetType();

            if (typeSource != typeDest) return;

            while (typeSource != typeof(object))
            {
                EventInfo[]? listeEvents = null;
                if (string.IsNullOrWhiteSpace(eventName))
                    listeEvents = typeSource.GetEvents(binding);
                else
                {
                    EventInfo ei = typeSource.GetEvent(eventName, binding);
                    if (ei != null)
                        listeEvents = new EventInfo[1] { ei };
                }

                if (listeEvents != null && listeEvents.Length > 0)
                {
                    foreach (EventInfo eventInfo in listeEvents)
                    {
                        string nomChamps = eventInfo.Name;
                        Delegate[]? delegues = GetAllEventHandlers(source, nomChamps);
                        if (delegues != null && delegues.Length > 0)
                            foreach (Delegate method in delegues)
                                eventInfo.AddEventHandler(dest, method);
                    }
                }
                typeSource = typeSource.BaseType;
            }
        }

        /// <summary>
        /// Copie les méthodes abonné aux events spécifiés (handlers) d'un objet à l'autre.
        /// Ainsi lorsqu'on recréer l'objet dest on récupère les handlers, pas besoin de se réabonner
        /// </summary>
        /// <param name="source">Objet source contenant les events abonnés</param>
        /// <param name="dest">Objet destination ou l'on va mettre les events abonné</param>
        /// <param name="listeEventName">Liste du/des events à copier</param>
        /// <remarks>Les deux objets (source et dest) doivent être obligatoirement du même type</remarks>
        public static void CopyEventsTo(object source, object dest, string[] listeEventName)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (dest == null) throw new ArgumentNullException(nameof(dest));

            foreach (string eventName in listeEventName)
            {
                CopyEventsTo(source, dest, eventName);
            }
        }

        /// <summary>
        /// Copie tous les events de l'objet vers un autre, inclus aussi tous les objets liés (en champs et/ou en propriétés)
        /// </summary>
        /// <param name="source">Instance de l'objet pour lequel on veut copier les events (et de tous les objets liés)</param>
        /// <param name="dest">Instance de l'objet pour lequel on va mettre les events</param>
        /// <param name="listeEvents">Liste du/des events à copier (ou null/vide pour tous les events)</param>
        public static void CopyEventsWithChildTo(object source, object dest, string[]? listeEvents = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (dest == null) throw new ArgumentNullException(nameof(dest));

            if (listeEvents == null)
                CopyEventsTo(source, dest);
            else
                CopyEventsTo(source, dest, listeEvents);
            foreach (FieldInfo fi in source.GetType().GetFields(binding))
                if (fi.GetValue(dest) != null)
                    CopyEventsWithChildTo(fi.GetValue(source), fi.GetValue(dest), listeEvents);
            foreach (PropertyInfo pi in source.GetType().GetProperties(binding))
                if (pi.GetValue(dest) != null)
                    CopyEventsWithChildTo(pi.GetValue(source), pi.GetValue(dest), listeEvents);
        }

        /// <summary>
        /// Copie tous les events des objets liés (exclu donc l'objet parent/racine, spécifié)
        /// </summary>
        /// <param name="source">Instance de l'objet parent/racine pour lequel on veut copier les events de tous les objets liés</param>
        /// <param name="dest">Instance parent/racine de l'objet pour lequel on va mettre les events</param>
        /// <param name="listeEvents">Liste du/des events à copier (ou null/vide pour tous les events)</param>
        public static void CopyChildEventsTo(object source, object dest, string[]? listeEvents = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (dest == null) throw new ArgumentNullException(nameof(dest));

            foreach (FieldInfo fi in source.GetType().GetFields(binding))
                if (fi.GetValue(dest) != null)
                    CopyEventsWithChildTo(fi.GetValue(source), fi.GetValue(dest), listeEvents);
            foreach (PropertyInfo pi in source.GetType().GetProperties(binding))
                if (pi.GetValue(dest) != null)
                    CopyEventsWithChildTo(pi.GetValue(source), pi.GetValue(dest), listeEvents);
        }

        /// <summary>
        /// Copie les events en utilisant le nom des méthodes, pour lier les events à leurs propres objets, inclus aussi tous les objets liés (en champs et/ou en propriétés)
        /// </summary>
        /// <param name="source">Instance de l'objet déclarant les events à copier</param>
        /// <param name="dest">Instance de l'objet ou l'on va recréer les events</param>
        /// <param name="listeEventsACopier">Liste des noms des events à recréer (ou tous si null/vide)</param>
        public static void CopyEventsWithChildByMethodNameTo(object source, object dest, string[]? listeEventsACopier = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (dest == null) throw new ArgumentNullException(nameof(dest));

            CopyEventsByMethodNameTo(source, dest, listeEventsACopier);
            foreach (FieldInfo fi in source.GetType().GetFields(binding))
                if (fi.GetValue(dest) != null)
                    CopyEventsWithChildByMethodNameTo(fi.GetValue(source), fi.GetValue(dest), listeEventsACopier);
            foreach (PropertyInfo pi in source.GetType().GetProperties(binding))
                if (pi.GetValue(dest) != null)
                    CopyEventsWithChildByMethodNameTo(pi.GetValue(source), pi.GetValue(dest), listeEventsACopier);
        }

        /// <summary>
        /// Copie tous les events des objets liés (exclu donc l'objet parent/racine, spécifié) en utilisant le nom des méthodes
        /// </summary>
        /// <param name="source">Instance de l'objet parent/racine pour lequel on veut copier les events de tous les objets liés</param>
        /// <param name="dest">Instance parent/racine de l'objet pour lequel on va mettre les events</param>
        /// <param name="listeEvents">Liste du/des events à copier (ou null/vide pour tous les events)</param>
        public static void CopyChildEventsByMethodNameTo(object source, object dest, string[]? listeEvents = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (dest == null) throw new ArgumentNullException(nameof(dest));

            foreach (FieldInfo fi in source.GetType().GetFields(binding))
                if (fi.GetValue(dest) != null)
                    CopyEventsWithChildByMethodNameTo(fi.GetValue(source), fi.GetValue(dest), listeEvents);
            foreach (PropertyInfo pi in source.GetType().GetProperties(binding))
                if (pi.GetValue(dest) != null)
                    CopyEventsWithChildByMethodNameTo(pi.GetValue(source), pi.GetValue(dest), listeEvents);
        }

        /// <summary>
        /// Copie les events en utilisant le nom des méthodes, pour lier les events à leurs propres objets
        /// </summary>
        /// <param name="source">Instance de l'objet déclarant les events à copier</param>
        /// <param name="dest">Instance de l'objet ou l'on va recréer les events</param>
        /// <param name="listeEventsACopier">Liste des noms des events à recréer (ou tous si null/vide)</param>
        public static void CopyEventsByMethodNameTo(object source, object dest, string[]? listeEventsACopier = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (dest == null) throw new ArgumentNullException(nameof(dest));

            List<string> dejaDeclare = new();

            Type typeSource = source.GetType();
            Type typeDest = dest.GetType();

            if (typeSource != typeDest) return;

            while (typeSource != typeof(object))
            {
                EventInfo[]? listeEvents = null;
                if (listeEventsACopier == null || listeEventsACopier.Length == 0)
                    listeEvents = typeSource.GetEvents(binding);
                else
                    listeEvents = typeSource.GetEvents(binding).Where(ei => listeEventsACopier.Contains(ei.Name.Trim(), StringComparer.OrdinalIgnoreCase)).ToArray();

                if (listeEvents != null && listeEvents.Length > 0)
                {
                    foreach (EventInfo eventInfo in listeEvents.Where(ei => !dejaDeclare.Contains(ei.DeclaringType + "|" + ei.Name)))
                    {
                        dejaDeclare.Add(eventInfo.DeclaringType + "|" + eventInfo.Name);
                        string nomChamps = eventInfo.Name;
                        Delegate[]? delegues = GetAllEventHandlers(source, nomChamps, out Type? delegue);
                        if (delegues != null && delegues.Length > 0)
                            foreach (MethodInfo method in delegues.Select(deleg => deleg.Method))
                            {
                                MethodInfo mi = method.DeclaringType.GetMethod(method.Name, binding, null, method.GetParameters().Select(pi => pi.ParameterType).ToArray(), null);
                                if (mi != null)
                                    eventInfo.AddEventHandler(dest, Delegate.CreateDelegate(eventInfo.EventHandlerType, mi.IsStatic ? null : dest, mi, true));
                            }
                    }
                }
                typeSource = typeSource.BaseType;
            }
        }

        #endregion

        #region Retourne délégué d'events

        /// <summary>
        /// Retourne tous les handlers abonné à l'évènement spécifié
        /// </summary>
        /// <param name="source">Objet source contenant l'event</param>
        /// <param name="nomEvent">Nom de l'event</param>
        /// <param name="typeDelegue">En retour, contient le type du délégué de cet event</param>
        public static Delegate[]? GetAllEventHandlers(object source, string nomEvent, out Type? typeDelegue)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrWhiteSpace(nomEvent)) throw new ArgumentNullException(nameof(nomEvent));

            Type typeSource = source.GetType();
            typeDelegue = null;

            string nomEventSystem = nomEvent;
            FieldInfo? champs = ChercheNomFieldEvent(typeSource, ref nomEvent);
            while (champs == null && typeSource != typeof(object))
            {
                typeSource = typeSource.BaseType;
                nomEvent = nomEventSystem;
                champs = ChercheNomFieldEvent(typeSource, ref nomEvent);
            }

            if (champs == null)
                return null;

            EventInfo ei = typeSource.GetEvent(nomEventSystem, binding);
            typeDelegue = ei.EventHandlerType;
            object retour = champs.GetValue(source);
            if (retour?.GetType() == typeof(object)) // Cas particulier event WinForm
            {
                PropertyInfo piEventHandlers = typeSource.GetProperty(NomProprieteEventWinForm, binding);
                EventHandlerList listeEventHandlers = (EventHandlerList)piEventHandlers.GetValue(source);
                return listeEventHandlers[retour]?.GetInvocationList();
            }
            else if (retour != null)
            {
                if (typeDelegue == typeof(RoutedEventHandler) || retour is RoutedEvent) // Cas particulier event WPF
                {
                    PropertyInfo piEventHandlersStore = typeSource.GetProperty(NomMagasinEventWpf, binding);
                    object eventHandlers = piEventHandlersStore.GetValue(source);
                    MethodInfo getRoutedEventHandlers = eventHandlers.GetType().GetMethod(NomMethodeRetourneHandlerMagasinEventWpf, binding);
                    RoutedEventHandlerInfo[] listeEventHandlers;
                    listeEventHandlers = (RoutedEventHandlerInfo[])getRoutedEventHandlers.Invoke(eventHandlers, new object[] { retour });
                    if (listeEventHandlers != null && listeEventHandlers.Length > 0)
                    {
                        Delegate[] delegates = new Delegate[0] { };
                        foreach (RoutedEventHandlerInfo routedEvent in listeEventHandlers)
                        {
                            Array.Resize(ref delegates, delegates.Length + 1);
                            delegates[delegates.Length - 1] = routedEvent.Handler;
                        }
                        return delegates;
                    }
                    else
                        return null;
                }
                else
                    return ((Delegate)Convert.ChangeType(retour, typeDelegue))?.GetInvocationList(); // Cas standard (event ajouté/codé soit même)
            }
            return null;
        }

        /// <summary>
        /// Retourne tous les handlers abonnés à l'évènement spécifié
        /// </summary>
        /// <param name="source">Objet source contenant l'event</param>
        /// <param name="nomEvent">Nom de l'event</param>
        public static Delegate[]? GetAllEventHandlers(object source, string nomEvent)
        {
            return GetAllEventHandlers(source, nomEvent, out _);
        }

        /// <summary>
        /// Retourne tous les handlers abonnés à tous les events de l'objet spécifié
        /// </summary>
        /// <param name="source">Objet contenant les events à rechercher</param>
        public static Dictionary<string, Delegate[]> GetAllEventsHandlers(object source)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            Dictionary<string, Delegate[]> retour = new();
            foreach (string nomEvent in source.GetType().GetEvents(binding).Select(ei => ei.Name))
            {
                Delegate[]? listeDelegues = GetAllEventHandlers(source, nomEvent);
                if (listeDelegues != null && listeDelegues.Length > 0)
                    retour.Add(nomEvent, listeDelegues);
            }
            return retour;
        }

        #endregion

        #region Supprime les délégués des events

        /// <summary>
        /// Supprime tous les délégués abonné à un event en particulier
        /// </summary>
        /// <param name="source">Instance de l'objet qui possède l'event à nettoyer</param>
        /// <param name="nomEvent">Nom de l'event à nettoyer</param>
        public static void RemoveAllEventHandlers(object source, string nomEvent)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (string.IsNullOrWhiteSpace(nomEvent)) throw new ArgumentNullException(nameof(nomEvent));

            Delegate[]? listeHandler = GetAllEventHandlers(source, nomEvent, out _);

            if (listeHandler == null || listeHandler.Length == 0) // Si y a pas d'abonné, il n'y a rien à supprimer, on sort tout de suite
                return;

            string nomChamps = nomEvent;

            Type typeSource = source.GetType();
            FieldInfo? champs = ChercheNomFieldEvent(source.GetType(), ref nomChamps);
            while (champs == null && typeSource != typeof(object))
            {
                nomChamps = nomEvent;
                typeSource = typeSource.BaseType;
                champs = ChercheNomFieldEvent(typeSource, ref nomChamps);
            }

            if (champs != null)
            {
                EventInfo ei = champs.DeclaringType.GetEvent(nomEvent, binding);
                foreach (Delegate d in listeHandler)
                {
                    ei.RemoveEventHandler(source, d);
                }
            }
        }

        /// <summary>
        /// Supprime tous les délégués de tous les events d'un objet spécifié
        /// </summary>
        /// <param name="source">Instance de l'objet possédant les events à nettoyer</param>
        /// <param name="listeEvents">Liste des events à nettoyer (ou null/vide pour tous les events)</param>
        public static void RemoveAllEventsHandlers(object source, string[]? listeEvents = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            EventInfo[] listeEventInfo;
            if (listeEvents == null || listeEvents.Length == 0)
                listeEventInfo = source.GetType().GetEvents(binding);
            else
                listeEventInfo = source.GetType().GetEvents(binding).Where(ei => listeEvents.Contains(ei.Name)).ToArray();
            foreach (string nomEvent in listeEventInfo.Select(ei => ei.Name))
            {
                RemoveAllEventHandlers(source, nomEvent);
            }
        }

        /// <summary>
        /// Supprime tous les délégués de tous les events d'un objet spécifié ainsi que ces objets liés (en champs ou propriétés)
        /// </summary>
        /// <param name="source">Instance de l'objet possédant les events à nettoyer</param>
        /// <param name="listeEvents">Liste des events à nettoyer (ou null/vide pour tous les events)</param>
        public static void RemoveAllEventsHandlersWithChild(object source, string[]? listeEvents = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            EventInfo[] listeEventInfo;
            if (listeEvents == null || listeEvents.Length == 0)
                listeEventInfo = source.GetType().GetEvents(binding);
            else
                listeEventInfo = source.GetType().GetEvents(binding).Where(ei => listeEvents.Contains(ei.Name)).ToArray();
            foreach (string nomEvent in listeEventInfo.Select(ei => ei.Name))
            {
                RemoveAllEventHandlers(source, nomEvent);
            }
            foreach (FieldInfo fi in source.GetType().GetFields(binding))
                if (fi.GetValue(source) != null)
                    RemoveAllEventsHandlersWithChild(fi.GetValue(source), listeEvents);
            foreach (PropertyInfo pi in source.GetType().GetProperties(binding))
                if (pi.GetValue(source) != null)
                    RemoveAllEventsHandlersWithChild(pi.GetValue(source), listeEvents);
        }

        /// <summary>
        /// Supprime tous les events des objets liés (exclu donc l'objet parent/racine, spécifié)
        /// </summary>
        /// <param name="source">Instance de l'objet parent/racine pour lequel on veut nettooyer les events de tous les objets liés</param>
        /// <param name="listeEvents">Liste du/des events à nettoyer (ou null/vide pour tous les events)</param>
        public static void RemoveChildEventsHandlers(object source, string[]? listeEvents = null)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            foreach (FieldInfo fi in source.GetType().GetFields(binding))
                if (fi.GetValue(source) != null)
                    RemoveAllEventsHandlersWithChild(fi.GetValue(source), listeEvents);
            foreach (PropertyInfo pi in source.GetType().GetProperties(binding))
                if (pi.GetValue(source) != null)
                    RemoveAllEventsHandlersWithChild(pi.GetValue(source), listeEvents);
        }

        #endregion

        /// <summary>
        /// Recherche le nom du champs stockant les handlers d'un event
        /// </summary>
        /// <param name="source">Objet contenant l'event à recherher</param>
        /// <param name="nomEvent">Nom de l'event pour lequel on veut le champs. Peut être modifié/retourné si le nom l'est pas le même (en particulier pour les Controls WinForm/WPF)</param>
        public static FieldInfo? ChercheNomFieldEvent(object source, ref string nomEvent)
        {
            return ChercheNomFieldEvent(source.GetType(), ref nomEvent);
        }

        /// <summary>
        /// Recherche le nom du champs stockant les handlers d'un event
        /// </summary>
        /// <param name="typeSource">Type de l'objet</param>
        /// <param name="nomEvent">Nom de l'event pour lequel on veut le champs. Peut être modifié/retourné si le nom l'est pas le même (en particulier pour les Controls WinForm/WPF)</param>
        public static FieldInfo? ChercheNomFieldEvent(Type typeSource, ref string nomEvent)
        {
            FieldInfo retour;
            retour = typeSource.GetField(nomEvent, binding);

            if (typeSource == typeof(Control) || typeSource.IsSubclassOf(typeof(Control)) ||
                typeSource == typeof(FrameworkElement) || typeSource.IsSubclassOf(typeof(FrameworkElement)))
            {
                if (retour == null)
                    retour = typeSource.GetField(NomChampsEventForm + nomEvent.Trim(), binding); // Cas des Form (en WinForm)
                if (retour == null)
                    retour = typeSource.GetField(NomChampsEventControlWinForm + nomEvent.Trim(), binding); // Cas des controls (en WinForm)
                if (retour == null)
                    retour = typeSource.GetField(nomEvent.Trim() + NomChampsEventControlWpf, binding); // Cas des controles (en WPF)
                if (retour == null)
                {
                    foreach (string suffixe in listeSuffixeNomChampsEventWinForm)
                    {
                        retour = typeSource.GetField(nomEvent.Trim() + suffixe, binding);
                        if (retour == null)
                            retour = typeSource.GetField(NomChampsEventForm + nomEvent.Trim() + suffixe, binding);
                        if (retour == null)
                            retour = typeSource.GetField(NomChampsEventControlWinForm + nomEvent.Trim() + suffixe, binding);
                        if (retour != null)
                            break;
                    }
                }
            }

            return retour;
        }

        /// <summary>
        /// Execute un BeginInvoke sur un event qui possède plusieurs abonnés (pas possible par BeginInvoke "classique", lève une exception)
        /// </summary>
        /// <param name="delegue">Délégué de l'évènement</param>
        /// <param name="source">Objet à l'origine de l'appel à l'évènement</param>
        /// <param name="nomEvent">Chaine de caractère du nom de l'évènement</param>
        /// <param name="parametres">Liste des paramètres</param>
        public static void BeginInvokeMultipleHandlers(this Delegate delegue, object source, string nomEvent, params object[] parametres)
        {
            if (delegue != null)
            {
                int nbParametres = delegue.GetType().GetMethod("Invoke").GetParameters().Count(pi => !pi.IsOptional);
                if (parametres.Length < nbParametres)
                    throw new Exception($"Pas assez de paramètres ; {nbParametres} minimum requis");
            }
            Delegate[]? retour = GetAllEventHandlers(source, nomEvent);
            if (retour != null && retour.Length > 0)
                foreach (Delegate handler in retour)
                    Task.Run(() =>
                    {
                        object[] param = new object[handler.Method.GetParameters().Length];
                        parametres.CopyTo(param, 0);
                        if (parametres.Length < handler.Method.GetParameters().Length)
                        {
                            for (int i = parametres.Length; i < handler.Method.GetParameters().Length; i++)
                                param[i] = handler.Method.GetParameters()[i].DefaultValue;
                        }
                        handler.Method.Invoke(handler.Target, param);
                    });
        }
    }
}

#nullable disable
