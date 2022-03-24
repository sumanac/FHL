import { useEffect, useState } from 'react';
import { RouteComponentProps } from 'react-router-dom';
import { Chat } from 'microsoft-graph';
import { getChat } from './GraphService';
import { useAppContext } from './AppContext';

import './Calendar.css';
import { useMsal } from '@azure/msal-react';
import { AccountInfo, InteractionRequiredAuthError, InteractionStatus } from '@azure/msal-browser';
import config from './Config';

export default function ChatDetail(props: RouteComponentProps) {
  const app = useAppContext();  

  const { instance, inProgress } = useMsal();
  const [chat, setChat] = useState<null|Chat>(null);

    useEffect(() => {
        if (!chat && inProgress === InteractionStatus.None) {
          getChat(app.authProvider!).then(response => setChat(response)).catch((e) => {
                if (e instanceof InteractionRequiredAuthError) {
                    instance.acquireTokenRedirect({
                        ...config,
                        account: instance.getActiveAccount() as AccountInfo
                    });
                }
            });
        }
    }, [app.authProvider, chat, inProgress, instance]);


  return (
    <div className="table-responsive">
      {chat &&
        <><div>
          Id: {chat.id}
        </div><div>
          Name: {chat.topic}
          </div><div>
          Chat type: {chat.chatType}
          </div></>}
    </div>
  );
}
