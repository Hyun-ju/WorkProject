using System;
using System.Diagnostics;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Ipc;

namespace SkHynix.HookUpToolkit.Exportor.IPC {
  public static class Utils {
    const string ServerName = "ExportorServer";
    const string ClientName = "ExportorClient";
    const string PortName = "ExportorPort";
    const string ObjectUri = "ExportorObject";


    public static string IpcUri() {
      return $"ipc://{PortName}/{ObjectUri}";
    }

    public static bool CreateServer() {
      try {
        var regChannels = ChannelServices.RegisteredChannels;
        var tmp = ChannelServices.GetChannel(ServerName);
        if (tmp != null) {
          Debug.WriteLine($"Has Server ==> {tmp.ChannelName}");
          return true;
        }

        var serverChannel = new IpcServerChannel(ServerName, PortName);
        ChannelServices.RegisterChannel(serverChannel, false);

        RemotingConfiguration.RegisterWellKnownServiceType(
            typeof(ExportorObject), ObjectUri, WellKnownObjectMode.Singleton);

        Debug.WriteLine($"Success to Create Server ==> {serverChannel.GetChannelUri()}");
      } catch (Exception ex) {
        Debug.WriteLine($"Fail to Create Server ==> {ex}");
        return false;
      }
      return true;
    }

    public static bool CreateClient() {
      try {
        var regChannels = ChannelServices.RegisteredChannels;
        var tmp = ChannelServices.GetChannel(ClientName);
        if (tmp != null) {
          Debug.WriteLine($"Has Server ==> {tmp.ChannelName}");
          return true;
        }

        var clientChannel = new IpcClientChannel(ClientName, null);
        ChannelServices.RegisterChannel(clientChannel, false);

        RemotingConfiguration.RegisterWellKnownClientType(
            typeof(ExportorObject), IpcUri());

        Debug.WriteLine($"Success to Create Client ==> {clientChannel.ChannelName}");
      } catch (Exception ex) {
        Debug.WriteLine($"Fail to Create Client ==> {ex}");
        return false;
      }
      return true;
    }

    public static bool UnregisterChennel() {
      try {
        var regChannels = ChannelServices.RegisteredChannels;
        var channel = ChannelServices.GetChannel(ClientName);
        if (channel == null) {
          return true;
        }
        ChannelServices.UnregisterChannel(channel);
      } catch (Exception ex) {
        return false;
      }
      return true;
    }

    public static ExportorObject GetShareObject() {
      return Activator.GetObject(typeof(ExportorObject), IpcUri()) as ExportorObject;
    }
  }
}
