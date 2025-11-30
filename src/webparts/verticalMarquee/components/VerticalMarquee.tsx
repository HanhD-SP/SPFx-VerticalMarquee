import * as React from 'react';
import styles from './VerticalMarquee.module.scss';
import type { IVerticalMarqueeProps } from './IVerticalMarqueeProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IListItem {
  Id: number;
  Title: string;
  [key: string]: any;
}

interface IVerticalMarqueeState {
  items: IListItem[];
  isLoading: boolean;
  error: string | null;
  isPaused: boolean;
}

export default class VerticalMarquee extends React.Component<IVerticalMarqueeProps, IVerticalMarqueeState> {
  private scrollWrapperRef: React.RefObject<HTMLDivElement>;
  private animationFrameId: number | null = null;
  private scrollPosition: number = 0;
  private scrollSpeed: number = 0.1;
  private lastTimestamp: number = 0;
  private itemSetHeight: number = 0;

  constructor(props: IVerticalMarqueeProps) {
    super(props);
    this.scrollWrapperRef = React.createRef<HTMLDivElement>();
    // Convert speed 1-10 to pixels per second (25% faster: divide by 12.8 instead of 16)
    const speed = props.scrollSpeed || 1;
    this.scrollSpeed = (speed / 12.8); // 25% faster: 1 = 0.078px/frame, 10 = 0.78px/frame
    this.state = {
      items: [],
      isLoading: true,
      error: null,
      isPaused: false
    };
  }

  public componentDidMount(): void {
    if (this.props.selectedList) {
      this.fetchListItems().then(() => {
        if (this.state.items.length > 0) {
          this.startScrolling();
        }
      });
    }
  }

  public componentDidUpdate(prevProps: IVerticalMarqueeProps): void {
    if (prevProps.selectedList !== this.props.selectedList) {
      this.stopScrolling();
      this.scrollPosition = 0;
      this.itemSetHeight = 0;
      if (this.scrollWrapperRef.current) {
        this.scrollWrapperRef.current.style.transform = 'translate3d(0, 0, 0)';
      }
      this.fetchListItems();
    }
    if (prevProps.scrollSpeed !== this.props.scrollSpeed) {
      const speed = typeof this.props.scrollSpeed === 'number' ? this.props.scrollSpeed : 1;
      this.scrollSpeed = speed / 12.8; // Convert 1-10 to 0.078-0.78 pixels per frame (25% faster)
    }
  }

  public componentWillUnmount(): void {
    this.stopScrolling();
  }

  private fetchListItems = async (): Promise<void> => {
    if (!this.props.selectedList || !this.props.context) {
      this.setState({ isLoading: false, items: [] });
      return;
    }

    this.setState({ isLoading: true, error: null });

    try {
      const listTitle = this.props.selectedList;
      const apiUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$select=Id,Title&$top=100`;

      const response: SPHttpClientResponse = await this.props.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Error fetching list items: ${response.statusText}`);
      }

      const data = await response.json();
      const items: IListItem[] = data.value || [];

      // Duplicate items for seamless scrolling - need at least 2 copies
      const duplicatedItems = items.length > 0 ? [...items, ...items] : [];

      this.setState({ items: duplicatedItems, isLoading: false, error: null }, () => {
        if (duplicatedItems.length > 0) {
          // Calculate item set height after render
          setTimeout(() => {
            if (this.scrollWrapperRef.current) {
              const firstItem = this.scrollWrapperRef.current.querySelector(`.${styles.marqueeItem}`) as HTMLElement;
              if (firstItem) {
                const itemHeight = firstItem.offsetHeight;
                this.itemSetHeight = itemHeight * items.length;
                this.scrollPosition = 0;
                this.startScrolling();
              }
            }
          }, 0);
        }
      });
    } catch (error) {
      this.setState({
        error: error instanceof Error ? error.message : 'Unknown error occurred',
        isLoading: false,
        items: []
      });
    }
  };

  private startScrolling = (): void => {
    this.lastTimestamp = performance.now();
    
    const scroll = (timestamp: number): void => {
      if (!this.state.isPaused && this.scrollWrapperRef.current && this.itemSetHeight > 0) {
        const wrapper = this.scrollWrapperRef.current;
        const deltaTime = timestamp - this.lastTimestamp;
        this.lastTimestamp = timestamp;
        
        // Smooth scrolling: use time-based movement for consistent speed
        // Frame rate independent - smooth on all devices
        const pixelsPerSecond = this.scrollSpeed * 60; // Convert to pixels per second
        const movement = (pixelsPerSecond * deltaTime) / 1000;
        
        this.scrollPosition += movement;

        // Seamless loop: reset position when reaching one complete set of items
        // Reset at exact item boundaries to eliminate any flash
        if (this.scrollPosition >= this.itemSetHeight) {
          this.scrollPosition = this.scrollPosition - this.itemSetHeight;
        }

        // Use transform with translate3d for better GPU acceleration
        // Apply transform to wrapper for smooth, seamless scrolling
        wrapper.style.transform = `translate3d(0, -${this.scrollPosition}px, 0)`;
        wrapper.style.willChange = 'transform';
      }

      this.animationFrameId = requestAnimationFrame(scroll);
    };

    this.animationFrameId = requestAnimationFrame(scroll);
  };

  private stopScrolling = (): void => {
    if (this.animationFrameId !== null) {
      cancelAnimationFrame(this.animationFrameId);
      this.animationFrameId = null;
    }
  };

  private handleMouseEnter = (): void => {
    this.setState({ isPaused: true });
  };

  private handleMouseLeave = (): void => {
    this.setState({ isPaused: false });
  };

  public render(): React.ReactElement<IVerticalMarqueeProps> {
    const { items, isLoading, error } = this.state;
    const { selectedList, textColor } = this.props;

    if (!selectedList) {
      return (
        <section className={styles.verticalMarquee}>
          <div className={styles.message}>
            Please select a list from the web part properties.
          </div>
        </section>
      );
    }

    if (isLoading) {
      return (
        <section className={styles.verticalMarquee}>
          <div className={styles.message}>Loading items...</div>
        </section>
      );
    }

    if (error) {
      return (
        <section className={styles.verticalMarquee}>
          <div className={styles.error}>Error: {error}</div>
        </section>
      );
    }

    if (items.length === 0) {
      return (
        <section className={styles.verticalMarquee}>
          <div className={styles.message}>No items found in the selected list.</div>
        </section>
      );
    }

    return (
      <section 
        className={styles.verticalMarquee}
        onMouseEnter={this.handleMouseEnter}
        onMouseLeave={this.handleMouseLeave}
      >
        <div className={styles.scrollContainer}>
          <div 
            ref={this.scrollWrapperRef}
            className={styles.scrollWrapper}
            style={{ color: textColor || '#000000' }}
          >
            {items.map((item, index) => (
              <div key={`${item.Id}-${index}`} className={styles.marqueeItem}>
                {item.Title || `Item ${item.Id}`}
              </div>
            ))}
          </div>
        </div>
      </section>
    );
  }
}
