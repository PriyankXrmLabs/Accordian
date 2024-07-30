import react, { useEffect ,useState} from 'react'
import { Accordion } from "@pnp/spfx-controls-react/lib/Accordion";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { getData } from './service';





export const Acc = (props) => {
    const [data, setData] = useState(null);
  
    useEffect(() => {
      const fetchData = async () => {
        try {
          const res = await getData(props.context, props.list);
          setData(res);
        } catch (error) {
          console.error('Error fetching data:', error);
        }
      };
  
      fetchData();
    }, [props.context, props.list]);
  
    return (
      <div>
       
        {data ? (
           data.map((item, index) => (
            <Accordion
              key={index}
              title={item.Title} // Assuming 'Title' is the field for accordion title
              defaultCollapsed={true}
              className="itemCell"
              collapsedIcon='ChevronDown'
              expandedIcon='ChevronUp'
            >
              <div style={{color:'black'}}>
              <RichText value={item.Description} isEditMode={false} />
              </div>
            </Accordion>
          ))
        ) : (
          <p>Loading...</p>
        )}
      </div>
    );
  };